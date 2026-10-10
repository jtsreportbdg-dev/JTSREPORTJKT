import React, { useState, useEffect, useMemo } from 'react';
import { AppState, StrategyMenuItem, DailyStrategyRecord } from '../types';
import { getStoredStrategyCatalog, getAreaForOutlet } from '../constants';
import { doc, onSnapshot, setDoc } from 'firebase/firestore';
import { db } from '../firebase';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

interface Props {
  data: AppState;
  filters: { month: string; year?: string; startDate: string; endDate?: string; outlet?: string };
  onUpdate: (updatedData: Partial<AppState>) => void;
  showToast?: (message: string, type?: 'success' | 'error' | 'info') => void;
  allMonthlyData?: Record<string, any>;
}

// Helper to compute COGS strategy target calculations dynamically
export const computeCogsStrategyTargets = (appData: AppState, targetDate: string) => {
  const monthKey = targetDate.substring(0, 7);
  const monthlyCogs = (appData.monthlyCogsData || {})[monthKey] || {};

  const parts = targetDate.split('-');
  const year = parseInt(parts[0], 10) || new Date().getFullYear();
  const month = parseInt(parts[1], 10) || 1;
  const day = parseInt(parts[2], 10) || 1;
  const daysInMonth = new Date(year, month, 0).getDate();

  // Target Revenue from System (strictly matching COGS feature logic)
  let strategyTargetRevenue = 0;
  if (appData?.dailyTargets?.[targetDate] !== undefined) {
    strategyTargetRevenue = Number(appData.dailyTargets[targetDate]) || 0;
  } else if (appData?.dailyTargets?.[String(day)] !== undefined) {
    strategyTargetRevenue = Number(appData.dailyTargets[String(day)]) || 0;
  } else if ((appData?.sales as any)?.dailyTargets?.[targetDate] !== undefined) {
    strategyTargetRevenue = Number((appData.sales as any).dailyTargets[targetDate]) || 0;
  } else if (appData?.sales?.dailyTarget) {
    strategyTargetRevenue = Number(appData.sales.dailyTarget) || 0;
  } else if (appData?.sales?.monthlyBudgets?.[monthKey] && daysInMonth > 0) {
    strategyTargetRevenue = Math.round(appData.sales.monthlyBudgets[monthKey] / daysInMonth);
  } else if (appData?.sales?.budgetRevenue && daysInMonth > 0) {
    strategyTargetRevenue = Math.round(appData.sales.budgetRevenue / daysInMonth);
  } else if (monthlyCogs.strategyTargetRevenue) {
    strategyTargetRevenue = Number(monthlyCogs.strategyTargetRevenue) || 0;
  }

  // Ratio from COGS (default 80% Percepatan / 20% Penyeimbang)
  const percepatanPercent = monthlyCogs.strategyPercepatanPercent !== undefined ? monthlyCogs.strategyPercepatanPercent : 80;
  const penyeimbangPercent = monthlyCogs.strategyPenyeimbangPercent !== undefined ? monthlyCogs.strategyPenyeimbangPercent : 20;

  // Contribution basis Lunch & Dinner from COGS
  const strategyBasis = monthlyCogs.strategyBasis || 'AUTO';
  let lunchContrib = 40;
  let dinnerContrib = 60;
  let basisName = 'Standar 40/60';

  if (strategyBasis === 'MANUAL') {
    lunchContrib = monthlyCogs.strategyManualLunchPercent !== undefined ? monthlyCogs.strategyManualLunchPercent : 40;
    dinnerContrib = monthlyCogs.strategyManualDinnerPercent !== undefined ? monthlyCogs.strategyManualDinnerPercent : 60;
    basisName = `Manual (${lunchContrib}% L / ${dinnerContrib}% D)`;
  } else {
    const rawRecordings = (appData as any).dailyRecording || appData.sales?.dailyRecording || {};
    const recordings = Object.entries(rawRecordings).filter(([d]) => d.startsWith(monthKey) && d <= targetDate);
    let sumLunch = 0;
    let sumDinner = 0;
    let daysWithData = 0;

    recordings.forEach(([_, rec]: [string, any]) => {
      const l = Number(rec?.lunchRevenue) || 0;
      const d = Number(rec?.dinnerRevenue) || 0;
      if (l > 0 || d > 0) {
        sumLunch += l;
        sumDinner += d;
        daysWithData++;
      }
    });

    if (daysWithData > 0 && (sumLunch + sumDinner) > 0) {
      lunchContrib = (sumLunch / (sumLunch + sumDinner)) * 100;
      dinnerContrib = (sumDinner / (sumLunch + sumDinner)) * 100;
      basisName = `MTD Aktual (${lunchContrib.toFixed(1)}% L / ${dinnerContrib.toFixed(1)}% D)`;
    }
  }

  // 1. Target Nominal per Session
  const targetLunchNominal = Math.round(strategyTargetRevenue * (lunchContrib / 100));
  const targetDinnerNominal = Math.round(strategyTargetRevenue * (dinnerContrib / 100));

  // 2. Target Sesuai Rasio COGS
  const targetLunchPercepatan = Math.round(targetLunchNominal * (percepatanPercent / 100));
  const targetLunchPenyeimbang = targetLunchNominal - targetLunchPercepatan;

  const targetDinnerPercepatan = Math.round(targetDinnerNominal * (percepatanPercent / 100));
  const targetDinnerPenyeimbang = targetDinnerNominal - targetDinnerPercepatan;

  return {
    strategyTargetRevenue,
    percepatanPercent,
    penyeimbangPercent,
    lunchContrib,
    dinnerContrib,
    basisName,
    targetLunchNominal,
    targetDinnerNominal,
    targetLunchPenyeimbang,
    targetLunchPercepatan,
    targetDinnerPenyeimbang,
    targetDinnerPercepatan,
  };
};

// Helper algorithm to distribute quantities so total nominal approaches the target
const autoDistributeQtyToApproachTarget = (
  selectedItemIds: string[],
  targetNominal: number,
  allCatalog: StrategyMenuItem[]
): Record<string, number> => {
  if (selectedItemIds.length === 0 || targetNominal <= 0) return {};

  const items = selectedItemIds
    .map(id => allCatalog.find(c => c.id === id))
    .filter((c): c is StrategyMenuItem => !!c && c.price > 0);

  if (items.length === 0) return {};

  items.sort((a, b) => b.price - a.price);

  const share = targetNominal / items.length;
  const result: Record<string, number> = {};
  let currentTotal = 0;

  items.forEach(it => {
    const qty = Math.max(1, Math.round(share / it.price));
    result[it.id] = qty;
    currentTotal += qty * it.price;
  });

  // Fine-tuning loop to minimize |currentTotal - targetNominal|
  let improved = true;
  let iterations = 0;
  while (improved && iterations < 30) {
    improved = false;
    iterations++;
    const diff = targetNominal - currentTotal;
    if (Math.abs(diff) < 2000) break;

    if (diff > 0) {
      let bestItem: StrategyMenuItem | null = null;
      let minGap = Math.abs(diff);
      for (const it of items) {
        const gap = Math.abs(diff - it.price);
        if (gap < minGap) {
          minGap = gap;
          bestItem = it;
        }
      }
      if (bestItem && minGap < Math.abs(diff)) {
        result[bestItem.id] = (result[bestItem.id] || 0) + 1;
        currentTotal += bestItem.price;
        improved = true;
      }
    } else {
      let bestItem: StrategyMenuItem | null = null;
      let minGap = Math.abs(diff);
      for (const it of items) {
        if ((result[it.id] || 0) > 1) {
          const gap = Math.abs(diff + it.price);
          if (gap < minGap) {
            minGap = gap;
            bestItem = it;
          }
        }
      }
      if (bestItem && minGap < Math.abs(diff)) {
        result[bestItem.id] = (result[bestItem.id] || 0) - 1;
        currentTotal -= bestItem.price;
        improved = true;
      }
    }
  }

  return result;
};

export const StrategiMenuModule: React.FC<Props> = ({
  data,
  filters,
  onUpdate,
  showToast
}) => {
  const [selectedDate, setSelectedDate] = useState<string>(filters.startDate || new Date().toISOString().split('T')[0]);
  const activeOutlet = filters.outlet || 'SH THE BARN';

  // Catalog State (Synced from system_config / Control Outlet)
  const [catalog, setCatalog] = useState<StrategyMenuItem[]>(() => getStoredStrategyCatalog());

  useEffect(() => {
    const handleCatalogEvent = (e: any) => {
      if (Array.isArray(e?.detail)) {
        setCatalog(e.detail);
      }
    };
    window.addEventListener('appStrategyCatalogChanged', handleCatalogEvent);

    const catRef = doc(db, 'system_config', 'strategy_catalog');
    const unsub = onSnapshot(catRef, (snap) => {
      if (snap.exists()) {
        const d = snap.data();
        if (Array.isArray(d.items) && d.items.length > 0) {
          setCatalog(d.items);
          try {
            localStorage.setItem('APP_STRATEGY_CATALOG', JSON.stringify(d.items));
          } catch (e) {}
        }
      }
    }, (err) => {
      console.warn('[StrategiMenuModule] Listener notice:', err);
    });

    return () => {
      window.removeEventListener('appStrategyCatalogChanged', handleCatalogEvent);
      unsub();
    };
  }, []);

  const activeCatalog = useMemo(() => {
    return catalog.filter(c => c.isActive);
  }, [catalog]);

  const penyeimbangCatalog = useMemo(() => {
    return activeCatalog.filter(c => c.category === 'PENYEIMBANG');
  }, [activeCatalog]);

  const percepatanCatalog = useMemo(() => {
    return activeCatalog.filter(c => c.category === 'PERCEPATAN');
  }, [activeCatalog]);

  // Current record for selected date
  const storedRecord: DailyStrategyRecord | null = useMemo(() => {
    return (data.dailyStrategy && data.dailyStrategy[selectedDate]) || null;
  }, [data.dailyStrategy, selectedDate]);

  // Compute COGS Strategy Targets dynamically for this date
  const cogsTargets = useMemo(() => {
    return computeCogsStrategyTargets(data, selectedDate);
  }, [data, selectedDate]);

  // Form State initialized from storedRecord or COGS auto-calculation
  const [targetPenyeimbangLunch, setTargetPenyeimbangLunch] = useState<number>(0);
  const [targetPercepatanLunch, setTargetPercepatanLunch] = useState<number>(0);
  const [targetPenyeimbangDinner, setTargetPenyeimbangDinner] = useState<number>(0);
  const [targetPercepatanDinner, setTargetPercepatanDinner] = useState<number>(0);

  // Selected item IDs
  const [selectedLunchPenyeimbang, setSelectedLunchPenyeimbang] = useState<string[]>([]);
  const [selectedLunchPercepatan, setSelectedLunchPercepatan] = useState<string[]>([]);
  const [selectedDinnerPenyeimbang, setSelectedDinnerPenyeimbang] = useState<string[]>([]);
  const [selectedDinnerPercepatan, setSelectedDinnerPercepatan] = useState<string[]>([]);

  // Item Qtys: { targetQty, actualQty }
  const [lunchPenyeimbangQtys, setLunchPenyeimbangQtys] = useState<Record<string, { targetQty: number; actualQty: number }>>({});
  const [lunchPercepatanQtys, setLunchPercepatanQtys] = useState<Record<string, { targetQty: number; actualQty: number }>>({});
  const [dinnerPenyeimbangQtys, setDinnerPenyeimbangQtys] = useState<Record<string, { targetQty: number; actualQty: number }>>({});
  const [dinnerPercepatanQtys, setDinnerPercepatanQtys] = useState<Record<string, { targetQty: number; actualQty: number }>>({});

  // Actual Revenue Kasir Inputs (Pulled from Daily Revenue / Sales module with override option)
  const todayDailyRecord = useMemo(() => {
    return (data.dailyRecording && data.dailyRecording[selectedDate]) || {};
  }, [data.dailyRecording, selectedDate]);

  const [actualLunchRevenueInput, setActualLunchRevenueInput] = useState<number>(0);
  const [actualDinnerRevenueInput, setActualDinnerRevenueInput] = useState<number>(0);
  const [isManualRevenueOverride, setIsManualRevenueOverride] = useState<boolean>(false);

  const [notes, setNotes] = useState<string>('');
  const [isSaving, setIsSaving] = useState(false);

  // Apply COGS targets helper
  const applyCogsTargets = () => {
    setTargetPenyeimbangLunch(cogsTargets.targetLunchPenyeimbang);
    setTargetPercepatanLunch(cogsTargets.targetLunchPercepatan);
    setTargetPenyeimbangDinner(cogsTargets.targetDinnerPenyeimbang);
    setTargetPercepatanDinner(cogsTargets.targetDinnerPercepatan);
    showToast?.('Target berhasil ditarik otomatis dari Fitur COGS!', 'success');
  };

  // Populate state when date or storedRecord changes
  useEffect(() => {
    if (storedRecord) {
      // If user had targets stored, load them; if both are 0, use COGS
      const hasStoredTargets = (storedRecord.targetPenyeimbangLunch || 0) > 0 || (storedRecord.targetPercepatanLunch || 0) > 0;
      setTargetPenyeimbangLunch(hasStoredTargets ? storedRecord.targetPenyeimbangLunch : cogsTargets.targetLunchPenyeimbang);
      setTargetPercepatanLunch(hasStoredTargets ? storedRecord.targetPercepatanLunch : cogsTargets.targetLunchPercepatan);
      setTargetPenyeimbangDinner(hasStoredTargets ? storedRecord.targetPenyeimbangDinner : cogsTargets.targetDinnerPenyeimbang);
      setTargetPercepatanDinner(hasStoredTargets ? storedRecord.targetPercepatanDinner : cogsTargets.targetDinnerPercepatan);

      setSelectedLunchPenyeimbang(storedRecord.selectedLunchPenyeimbang || penyeimbangCatalog.slice(0, 2).map(c => c.id));
      setSelectedLunchPercepatan(storedRecord.selectedLunchPercepatan || percepatanCatalog.slice(0, 2).map(c => c.id));
      setSelectedDinnerPenyeimbang(storedRecord.selectedDinnerPenyeimbang || penyeimbangCatalog.slice(0, 2).map(c => c.id));
      setSelectedDinnerPercepatan(storedRecord.selectedDinnerPercepatan || percepatanCatalog.slice(0, 2).map(c => c.id));

      const lPen: Record<string, { targetQty: number; actualQty: number }> = {};
      Object.entries(storedRecord.lunchPenyeimbangItems || {}).forEach(([id, it]) => {
        lPen[id] = { targetQty: it.targetQty || 0, actualQty: it.actualQty || 0 };
      });
      setLunchPenyeimbangQtys(lPen);

      const lPer: Record<string, { targetQty: number; actualQty: number }> = {};
      Object.entries(storedRecord.lunchPercepatanItems || {}).forEach(([id, it]) => {
        lPer[id] = { targetQty: it.targetQty || 0, actualQty: it.actualQty || 0 };
      });
      setLunchPercepatanQtys(lPer);

      const dPen: Record<string, { targetQty: number; actualQty: number }> = {};
      Object.entries(storedRecord.dinnerPenyeimbangItems || {}).forEach(([id, it]) => {
        dPen[id] = { targetQty: it.targetQty || 0, actualQty: it.actualQty || 0 };
      });
      setDinnerPenyeimbangQtys(dPen);

      const dPer: Record<string, { targetQty: number; actualQty: number }> = {};
      Object.entries(storedRecord.dinnerPercepatanItems || {}).forEach(([id, it]) => {
        dPer[id] = { targetQty: it.targetQty || 0, actualQty: it.actualQty || 0 };
      });
      setDinnerPercepatanQtys(dPer);

      setNotes(storedRecord.notes || '');

      // Revenue inputs
      const sysLunch = Number(todayDailyRecord.lunchRevenue) || 0;
      const sysDinner = Number(todayDailyRecord.dinnerRevenue) || 0;
      if (storedRecord.actualRevenueLunchTotal && storedRecord.actualRevenueLunchTotal > 0 && sysLunch === 0) {
        setActualLunchRevenueInput(storedRecord.actualRevenueLunchTotal);
        setActualDinnerRevenueInput(storedRecord.actualRevenueDinnerTotal || 0);
        setIsManualRevenueOverride(true);
      } else {
        setActualLunchRevenueInput(sysLunch);
        setActualDinnerRevenueInput(sysDinner);
        setIsManualRevenueOverride(false);
      }
    } else {
      // Auto from COGS by default!
      setTargetPenyeimbangLunch(cogsTargets.targetLunchPenyeimbang);
      setTargetPercepatanLunch(cogsTargets.targetLunchPercepatan);
      setTargetPenyeimbangDinner(cogsTargets.targetDinnerPenyeimbang);
      setTargetPercepatanDinner(cogsTargets.targetDinnerPercepatan);

      const defaultPenIds = penyeimbangCatalog.slice(0, 2).map(c => c.id);
      const defaultPerIds = percepatanCatalog.slice(0, 2).map(c => c.id);

      setSelectedLunchPenyeimbang(defaultPenIds);
      setSelectedLunchPercepatan(defaultPerIds);
      setSelectedDinnerPenyeimbang(defaultPenIds);
      setSelectedDinnerPercepatan(defaultPerIds);

      // Auto-set initial target quantities to approach COGS target!
      const initLPen = autoDistributeQtyToApproachTarget(defaultPenIds, cogsTargets.targetLunchPenyeimbang, catalog);
      const initLPer = autoDistributeQtyToApproachTarget(defaultPerIds, cogsTargets.targetLunchPercepatan, catalog);
      const initDPen = autoDistributeQtyToApproachTarget(defaultPenIds, cogsTargets.targetDinnerPenyeimbang, catalog);
      const initDPer = autoDistributeQtyToApproachTarget(defaultPerIds, cogsTargets.targetDinnerPercepatan, catalog);

      const lPen: Record<string, { targetQty: number; actualQty: number }> = {};
      Object.entries(initLPen).forEach(([id, qty]) => { lPen[id] = { targetQty: qty, actualQty: 0 }; });
      setLunchPenyeimbangQtys(lPen);

      const lPer: Record<string, { targetQty: number; actualQty: number }> = {};
      Object.entries(initLPer).forEach(([id, qty]) => { lPer[id] = { targetQty: qty, actualQty: 0 }; });
      setLunchPercepatanQtys(lPer);

      const dPen: Record<string, { targetQty: number; actualQty: number }> = {};
      Object.entries(initDPen).forEach(([id, qty]) => { dPen[id] = { targetQty: qty, actualQty: 0 }; });
      setDinnerPenyeimbangQtys(dPen);

      const dPer: Record<string, { targetQty: number; actualQty: number }> = {};
      Object.entries(initDPer).forEach(([id, qty]) => { dPer[id] = { targetQty: qty, actualQty: 0 }; });
      setDinnerPercepatanQtys(dPer);

      setNotes('');

      const sysLunch = Number(todayDailyRecord.lunchRevenue) || 0;
      const sysDinner = Number(todayDailyRecord.dinnerRevenue) || 0;
      setActualLunchRevenueInput(sysLunch);
      setActualDinnerRevenueInput(sysDinner);
      setIsManualRevenueOverride(false);
    }
  }, [storedRecord, selectedDate, cogsTargets, todayDailyRecord]);

  // Helper to get catalog item
  const getItemDetails = (id: string): StrategyMenuItem => {
    const found = catalog.find(c => c.id === id);
    if (found) return found;
    return {
      id,
      name: `Menu #${id}`,
      category: 'PENYEIMBANG',
      price: 0,
      isActive: true
    };
  };

  // Calculations for Lunch Penyeimbang
  const lunchPenyeimbangRows = useMemo(() => {
    return selectedLunchPenyeimbang.map(id => {
      const item = getItemDetails(id);
      const qtyData = lunchPenyeimbangQtys[id] || { targetQty: 0, actualQty: 0 };
      const targetRev = qtyData.targetQty * item.price;
      const actualRev = qtyData.actualQty * item.price;
      const gapRev = actualRev - targetRev;
      const pct = targetRev > 0 ? (actualRev / targetRev) * 100 : actualRev > 0 ? 100 : 0;
      return { ...item, targetQty: qtyData.targetQty, actualQty: qtyData.actualQty, targetRev, actualRev, gapRev, pct };
    });
  }, [selectedLunchPenyeimbang, lunchPenyeimbangQtys, catalog]);

  // Calculations for Lunch Percepatan
  const lunchPercepatanRows = useMemo(() => {
    return selectedLunchPercepatan.map(id => {
      const item = getItemDetails(id);
      const qtyData = lunchPercepatanQtys[id] || { targetQty: 0, actualQty: 0 };
      const targetRev = qtyData.targetQty * item.price;
      const actualRev = qtyData.actualQty * item.price;
      const gapRev = actualRev - targetRev;
      const pct = targetRev > 0 ? (actualRev / targetRev) * 100 : actualRev > 0 ? 100 : 0;
      return { ...item, targetQty: qtyData.targetQty, actualQty: qtyData.actualQty, targetRev, actualRev, gapRev, pct };
    });
  }, [selectedLunchPercepatan, lunchPercepatanQtys, catalog]);

  // Calculations for Dinner Penyeimbang
  const dinnerPenyeimbangRows = useMemo(() => {
    return selectedDinnerPenyeimbang.map(id => {
      const item = getItemDetails(id);
      const qtyData = dinnerPenyeimbangQtys[id] || { targetQty: 0, actualQty: 0 };
      const targetRev = qtyData.targetQty * item.price;
      const actualRev = qtyData.actualQty * item.price;
      const gapRev = actualRev - targetRev;
      const pct = targetRev > 0 ? (actualRev / targetRev) * 100 : actualRev > 0 ? 100 : 0;
      return { ...item, targetQty: qtyData.targetQty, actualQty: qtyData.actualQty, targetRev, actualRev, gapRev, pct };
    });
  }, [selectedDinnerPenyeimbang, dinnerPenyeimbangQtys, catalog]);

  // Calculations for Dinner Percepatan
  const dinnerPercepatanRows = useMemo(() => {
    return selectedDinnerPercepatan.map(id => {
      const item = getItemDetails(id);
      const qtyData = dinnerPercepatanQtys[id] || { targetQty: 0, actualQty: 0 };
      const targetRev = qtyData.targetQty * item.price;
      const actualRev = qtyData.actualQty * item.price;
      const gapRev = actualRev - targetRev;
      const pct = targetRev > 0 ? (actualRev / targetRev) * 100 : actualRev > 0 ? 100 : 0;
      return { ...item, targetQty: qtyData.targetQty, actualQty: qtyData.actualQty, targetRev, actualRev, gapRev, pct };
    });
  }, [selectedDinnerPercepatan, dinnerPercepatanQtys, catalog]);

  // Planned Menu Target Totals (Accumulated nominal of menus approaching target)
  const plannedTargetLunchPenyeimbang = useMemo(() => lunchPenyeimbangRows.reduce((a, b) => a + b.targetRev, 0), [lunchPenyeimbangRows]);
  const plannedTargetLunchPercepatan = useMemo(() => lunchPercepatanRows.reduce((a, b) => a + b.targetRev, 0), [lunchPercepatanRows]);
  const plannedTargetDinnerPenyeimbang = useMemo(() => dinnerPenyeimbangRows.reduce((a, b) => a + b.targetRev, 0), [dinnerPenyeimbangRows]);
  const plannedTargetDinnerPercepatan = useMemo(() => dinnerPercepatanRows.reduce((a, b) => a + b.targetRev, 0), [dinnerPercepatanRows]);

  // Actual Menu Revenue Totals
  const actualRevLunchPenyeimbang = useMemo(() => lunchPenyeimbangRows.reduce((a, b) => a + b.actualRev, 0), [lunchPenyeimbangRows]);
  const actualRevLunchPercepatan = useMemo(() => lunchPercepatanRows.reduce((a, b) => a + b.actualRev, 0), [lunchPercepatanRows]);
  const actualRevDinnerPenyeimbang = useMemo(() => dinnerPenyeimbangRows.reduce((a, b) => a + b.actualRev, 0), [dinnerPenyeimbangRows]);
  const actualRevDinnerPercepatan = useMemo(() => dinnerPercepatanRows.reduce((a, b) => a + b.actualRev, 0), [dinnerPercepatanRows]);

  const totalTargetLunch = targetPenyeimbangLunch + targetPercepatanLunch;
  const totalActualLunchMenu = actualRevLunchPenyeimbang + actualRevLunchPercepatan;
  const pctLunch = totalTargetLunch > 0 ? (totalActualLunchMenu / totalTargetLunch) * 100 : 0;

  const totalTargetDinner = targetPenyeimbangDinner + targetPercepatanDinner;
  const totalActualDinnerMenu = actualRevDinnerPenyeimbang + actualRevDinnerPercepatan;
  const pctDinner = totalTargetDinner > 0 ? (totalActualDinnerMenu / totalTargetDinner) * 100 : 0;

  const grandTotalTarget = totalTargetLunch + totalTargetDinner;
  const grandTotalActualMenu = totalActualLunchMenu + totalActualDinnerMenu;
  const grandPct = grandTotalTarget > 0 ? (grandTotalActualMenu / grandTotalTarget) * 100 : 0;

  const totalTargetPenyeimbang = targetPenyeimbangLunch + targetPenyeimbangDinner;
  const totalActualPenyeimbang = actualRevLunchPenyeimbang + actualRevDinnerPenyeimbang;
  const pctPenyeimbang = totalTargetPenyeimbang > 0 ? (totalActualPenyeimbang / totalTargetPenyeimbang) * 100 : 0;

  const totalTargetPercepatan = targetPercepatanLunch + targetPercepatanDinner;
  const totalActualPercepatan = actualRevLunchPercepatan + actualRevDinnerPercepatan;
  const pctPercepatan = totalTargetPercepatan > 0 ? (totalActualPercepatan / totalTargetPercepatan) * 100 : 0;

  // Actual Revenue Kasir vs Outside Menu Analysis ("DI LUAR PENJUALAN MENU")
  const effectiveLunchRevenue = isManualRevenueOverride ? actualLunchRevenueInput : (Number(todayDailyRecord.lunchRevenue) || actualLunchRevenueInput);
  const effectiveDinnerRevenue = isManualRevenueOverride ? actualDinnerRevenueInput : (Number(todayDailyRecord.dinnerRevenue) || actualDinnerRevenueInput);
  const effectiveTotalRevenue = effectiveLunchRevenue + effectiveDinnerRevenue;

  const actualLunchOutsideMenu = Math.max(0, effectiveLunchRevenue - totalActualLunchMenu);
  const actualDinnerOutsideMenu = Math.max(0, effectiveDinnerRevenue - totalActualDinnerMenu);
  const actualGrandOutsideMenu = Math.max(0, effectiveTotalRevenue - grandTotalActualMenu);

  const contribLunchMenuPct = effectiveLunchRevenue > 0 ? (totalActualLunchMenu / effectiveLunchRevenue) * 100 : 0;
  const contribLunchOutsidePct = effectiveLunchRevenue > 0 ? (actualLunchOutsideMenu / effectiveLunchRevenue) * 100 : 0;

  const contribDinnerMenuPct = effectiveDinnerRevenue > 0 ? (totalActualDinnerMenu / effectiveDinnerRevenue) * 100 : 0;
  const contribDinnerOutsidePct = effectiveDinnerRevenue > 0 ? (actualDinnerOutsideMenu / effectiveDinnerRevenue) * 100 : 0;

  const contribGrandMenuPct = effectiveTotalRevenue > 0 ? (grandTotalActualMenu / effectiveTotalRevenue) * 100 : 0;
  const contribGrandOutsidePct = effectiveTotalRevenue > 0 ? (actualGrandOutsideMenu / effectiveTotalRevenue) * 100 : 0;

  // Auto-set menu quantities helper to approach segment target
  const handleAutoSetMenu = (
    segment: 'LUNCH_PENYEIMBANG' | 'LUNCH_PERCEPATAN' | 'DINNER_PENYEIMBANG' | 'DINNER_PERCEPATAN'
  ) => {
    let ids: string[] = [];
    let target = 0;
    let setter: React.Dispatch<React.SetStateAction<Record<string, { targetQty: number; actualQty: number }>>>;

    if (segment === 'LUNCH_PENYEIMBANG') {
      ids = selectedLunchPenyeimbang;
      target = targetPenyeimbangLunch;
      setter = setLunchPenyeimbangQtys;
    } else if (segment === 'LUNCH_PERCEPATAN') {
      ids = selectedLunchPercepatan;
      target = targetPercepatanLunch;
      setter = setLunchPercepatanQtys;
    } else if (segment === 'DINNER_PENYEIMBANG') {
      ids = selectedDinnerPenyeimbang;
      target = targetPenyeimbangDinner;
      setter = setDinnerPenyeimbangQtys;
    } else {
      ids = selectedDinnerPercepatan;
      target = targetPercepatanDinner;
      setter = setDinnerPercepatanQtys;
    }

    if (ids.length === 0) {
      showToast?.('Pilih menu terlebih dahulu untuk segmen ini!', 'info');
      return;
    }
    if (target <= 0) {
      showToast?.('Target nominal belum ditentukan atau bernilai 0!', 'info');
      return;
    }

    const calculated = autoDistributeQtyToApproachTarget(ids, target, catalog);
    setter(prev => {
      const next = { ...prev };
      Object.entries(calculated).forEach(([id, qty]) => {
        next[id] = { ...(next[id] || { actualQty: 0 }), targetQty: qty };
      });
      return next;
    });

    showToast?.('Kuantitas menu berhasil diatur mendekati nominal target!', 'success');
  };

  // Toggle item selection helper
  const toggleSelection = (
    currentList: string[],
    setList: React.Dispatch<React.SetStateAction<string[]>>,
    itemId: string
  ) => {
    if (currentList.includes(itemId)) {
      setList(currentList.filter(id => id !== itemId));
    } else {
      setList([...currentList, itemId]);
    }
  };

  // Save handler
  const handleSave = async () => {
    setIsSaving(true);
    try {
      const lunchPenItems: Record<string, any> = {};
      lunchPenyeimbangRows.forEach(r => {
        lunchPenItems[r.id] = {
          targetQty: r.targetQty,
          actualQty: r.actualQty,
          targetRevenue: r.targetRev,
          actualRevenue: r.actualRev
        };
      });

      const lunchPerItems: Record<string, any> = {};
      lunchPercepatanRows.forEach(r => {
        lunchPerItems[r.id] = {
          targetQty: r.targetQty,
          actualQty: r.actualQty,
          targetRevenue: r.targetRev,
          actualRevenue: r.actualRev
        };
      });

      const dinnerPenItems: Record<string, any> = {};
      dinnerPenyeimbangRows.forEach(r => {
        dinnerPenItems[r.id] = {
          targetQty: r.targetQty,
          actualQty: r.actualQty,
          targetRevenue: r.targetRev,
          actualRevenue: r.actualRev
        };
      });

      const dinnerPerItems: Record<string, any> = {};
      dinnerPercepatanRows.forEach(r => {
        dinnerPerItems[r.id] = {
          targetQty: r.targetQty,
          actualQty: r.actualQty,
          targetRevenue: r.targetRev,
          actualRevenue: r.actualRev
        };
      });

      const newRecord: DailyStrategyRecord = {
        targetPenyeimbangLunch,
        targetPercepatanLunch,
        targetPenyeimbangDinner,
        targetPercepatanDinner,
        totalTargetLunch,
        totalTargetDinner,
        grandTotalTarget,

        selectedLunchPenyeimbang,
        selectedLunchPercepatan,
        selectedDinnerPenyeimbang,
        selectedDinnerPercepatan,

        lunchPenyeimbangItems: lunchPenItems,
        lunchPercepatanItems: lunchPerItems,
        dinnerPenyeimbangItems: dinnerPenItems,
        dinnerPercepatanItems: dinnerPerItems,

        actualRevenuePenyeimbangLunch: actualRevLunchPenyeimbang,
        actualRevenuePercepatanLunch: actualRevLunchPercepatan,
        actualRevenuePenyeimbangDinner: actualRevDinnerPenyeimbang,
        actualRevenuePercepatanDinner: actualRevDinnerPercepatan,

        totalActualLunch: totalActualLunchMenu,
        totalActualDinner: totalActualDinnerMenu,
        grandTotalActual: grandTotalActualMenu,

        // Actual Revenue Total & Di Luar Penjualan Menu
        actualRevenueLunchTotal: effectiveLunchRevenue,
        actualRevenueDinnerTotal: effectiveDinnerRevenue,
        actualRevenueGrandTotal: effectiveTotalRevenue,
        actualRevenueLunchOutsideMenu: actualLunchOutsideMenu,
        actualRevenueDinnerOutsideMenu: actualDinnerOutsideMenu,
        actualRevenueGrandOutsideMenu: actualGrandOutsideMenu,
        contributionLunchMenuPercent: contribLunchMenuPct,
        contributionDinnerMenuPercent: contribDinnerMenuPct,
        contributionGrandMenuPercent: contribGrandMenuPct,
        contributionLunchOutsidePercent: contribLunchOutsidePct,
        contributionDinnerOutsidePercent: contribDinnerOutsidePct,
        contributionGrandOutsidePercent: contribGrandOutsidePct,

        // Auto COGS tracking
        isTargetAutoFromCogs: true,
        cogsTargetRevenue: cogsTargets.strategyTargetRevenue,
        cogsPercepatanPercent: cogsTargets.percepatanPercent,
        cogsPenyeimbangPercent: cogsTargets.penyeimbangPercent,

        notes,
        updatedAt: new Date().toISOString()
      };

      const updatedDailyStrategy = {
        ...(data.dailyStrategy || {}),
        [selectedDate]: newRecord
      };

      onUpdate({ dailyStrategy: updatedDailyStrategy });

      try {
        const outletDocRef = doc(db, 'dashboards', activeOutlet);
        await setDoc(outletDocRef, {
          dailyStrategy: updatedDailyStrategy,
          updatedAt: new Date().toISOString()
        }, { merge: true });
      } catch (err) {}

      showToast?.(`Data Strategi & Analisa Kontribusi tanggal ${selectedDate} berhasil disimpan!`, 'success');
    } catch (e) {
      console.error(e);
      showToast?.('Gagal menyimpan data strategi. Silakan coba lagi.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  // WhatsApp formatted copy
  const handleCopyWhatsApp = () => {
    const area = getAreaForOutlet(activeOutlet);
    const text = `*JUSTUS STEAKHOUSE - LAPORAN MENU STRATEGI & KONTRIBUSI REVENUE*

*OUTLET* : ${activeOutlet} (${area})
*TANGGAL* : ${selectedDate}

*1. TARGET OTOMATIS DARI FITUR COGS*
• Target Today Revenue : Rp ${cogsTargets.strategyTargetRevenue.toLocaleString('id-ID')}
• Rasio Strategi : *${cogsTargets.percepatanPercent}% Percepatan* | *${cogsTargets.penyeimbangPercent}% Penyeimbang*
• Basis Kontribusi : ${cogsTargets.basisName}
• Target Lunch : Rp ${totalTargetLunch.toLocaleString('id-ID')} (Penyeimbang: Rp ${targetPenyeimbangLunch.toLocaleString('id-ID')} | Percepatan: Rp ${targetPercepatanLunch.toLocaleString('id-ID')})
• Target Dinner : Rp ${totalTargetDinner.toLocaleString('id-ID')} (Penyeimbang: Rp ${targetPenyeimbangDinner.toLocaleString('id-ID')} | Percepatan: Rp ${targetPercepatanDinner.toLocaleString('id-ID')})
• *Grand Target Strategi* : *Rp ${grandTotalTarget.toLocaleString('id-ID')}*

*2. PENCAPAIAN ACTUAL PENJUALAN MENU STRATEGI*
• Lunch Penyeimbang : Rp ${actualRevLunchPenyeimbang.toLocaleString('id-ID')}
• Lunch Percepatan : Rp ${actualRevLunchPercepatan.toLocaleString('id-ID')}
• *Total Actual Menu Lunch* : Rp ${totalActualLunchMenu.toLocaleString('id-ID')} (${pctLunch.toFixed(1)}% vs Target)
• Dinner Penyeimbang : Rp ${actualRevDinnerPenyeimbang.toLocaleString('id-ID')}
• Dinner Percepatan : Rp ${actualRevDinnerPercepatan.toLocaleString('id-ID')}
• *Total Actual Menu Dinner* : Rp ${totalActualDinnerMenu.toLocaleString('id-ID')} (${pctDinner.toFixed(1)}% vs Target)
• *GRAND TOTAL PENJUALAN MENU* : *Rp ${grandTotalActualMenu.toLocaleString('id-ID')}* (${grandPct.toFixed(1)}%)

*3. ANALISA KONTRIBUSI TERHADAP ACTUAL REVENUE (DI LUAR PENJUALAN MENU)*
☀️ *SHIFT LUNCH*:
- Actual Revenue Kasir : Rp ${effectiveLunchRevenue.toLocaleString('id-ID')}
- Penjualan Menu Strategi : Rp ${totalActualLunchMenu.toLocaleString('id-ID')} (*${contribLunchMenuPct.toFixed(1)}%*)
- Revenue Di Luar Menu : Rp ${actualLunchOutsideMenu.toLocaleString('id-ID')} (*${contribLunchOutsidePct.toFixed(1)}%*)

🌙 *SHIFT DINNER*:
- Actual Revenue Kasir : Rp ${effectiveDinnerRevenue.toLocaleString('id-ID')}
- Penjualan Menu Strategi : Rp ${totalActualDinnerMenu.toLocaleString('id-ID')} (*${contribDinnerMenuPct.toFixed(1)}%*)
- Revenue Di Luar Menu : Rp ${actualDinnerOutsideMenu.toLocaleString('id-ID')} (*${contribDinnerOutsidePct.toFixed(1)}%*)

👑 *TOTAL HARIAN (FULL DAY)*:
- Total Actual Revenue Kasir : *Rp ${effectiveTotalRevenue.toLocaleString('id-ID')}*
- Total Penjualan Menu Strategi : *Rp ${grandTotalActualMenu.toLocaleString('id-ID')}* (*${contribGrandMenuPct.toFixed(1)}%*)
- Total Di Luar Penjualan Menu : *Rp ${actualGrandOutsideMenu.toLocaleString('id-ID')}* (*${contribGrandOutsidePct.toFixed(1)}%*)

*Catatan / Evaluasi*:
${notes || '-'}`;

    navigator.clipboard.writeText(text);
    showToast?.('Format WhatsApp berhasil disalin ke clipboard!', 'success');
  };

  // Export to Excel
  const handleExportExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet(`Strategi_${selectedDate}`);

      sheet.addRow([`JUSTUS STEAKHOUSE - LAPORAN MENU STRATEGI & ANALISA KONTRIBUSI`]);
      sheet.addRow([`Outlet: ${activeOutlet} | Tanggal: ${selectedDate}`]);
      sheet.addRow([`Target COGS: Rp ${cogsTargets.strategyTargetRevenue.toLocaleString('id-ID')} (${cogsTargets.percepatanPercent}% Percepatan / ${cogsTargets.penyeimbangPercent}% Penyeimbang)`]);
      sheet.addRow([]);

      sheet.addRow(['1. ANALISA KONTRIBUSI REVENUE (PENJUALAN MENU VS DI LUAR MENU)']);
      sheet.addRow(['Sesi', 'Actual Revenue Kasir (Rp)', 'Penjualan Menu Strategi (Rp)', '% Kontribusi Menu', 'Revenue Di Luar Menu (Rp)', '% Kontribusi Luar Menu']);
      sheet.addRow(['Lunch', effectiveLunchRevenue, totalActualLunchMenu, `${contribLunchMenuPct.toFixed(1)}%`, actualLunchOutsideMenu, `${contribLunchOutsidePct.toFixed(1)}%`]);
      sheet.addRow(['Dinner', effectiveDinnerRevenue, totalActualDinnerMenu, `${contribDinnerMenuPct.toFixed(1)}%`, actualDinnerOutsideMenu, `${contribDinnerOutsidePct.toFixed(1)}%`]);
      sheet.addRow(['TOTAL', effectiveTotalRevenue, grandTotalActualMenu, `${contribGrandMenuPct.toFixed(1)}%`, actualGrandOutsideMenu, `${contribGrandOutsidePct.toFixed(1)}%`]);
      sheet.addRow([]);

      sheet.addRow(['2. RINGKASAN TARGET & ACTUAL STRATEGI MENU']);
      sheet.addRow(['Shift', 'Kategori', 'Target Revenue (Rp)', 'Planned Menu (Rp)', 'Actual Revenue (Rp)', 'Selisih (Rp)', 'Pencapaian (%)']);
      sheet.addRow(['Lunch', 'Penyeimbang', targetPenyeimbangLunch, plannedTargetLunchPenyeimbang, actualRevLunchPenyeimbang, actualRevLunchPenyeimbang - targetPenyeimbangLunch, targetPenyeimbangLunch > 0 ? ((actualRevLunchPenyeimbang / targetPenyeimbangLunch) * 100).toFixed(1) + '%' : '-']);
      sheet.addRow(['Lunch', 'Percepatan', targetPercepatanLunch, plannedTargetLunchPercepatan, actualRevLunchPercepatan, actualRevLunchPercepatan - targetPercepatanLunch, targetPercepatanLunch > 0 ? ((actualRevLunchPercepatan / targetPercepatanLunch) * 100).toFixed(1) + '%' : '-']);
      sheet.addRow(['Dinner', 'Penyeimbang', targetPenyeimbangDinner, plannedTargetDinnerPenyeimbang, actualRevDinnerPenyeimbang, actualRevDinnerPenyeimbang - targetPenyeimbangDinner, targetPenyeimbangDinner > 0 ? ((actualRevDinnerPenyeimbang / targetPenyeimbangDinner) * 100).toFixed(1) + '%' : '-']);
      sheet.addRow(['Dinner', 'Percepatan', targetPercepatanDinner, plannedTargetDinnerPercepatan, actualRevDinnerPercepatan, actualRevDinnerPercepatan - targetPercepatanDinner, targetPercepatanDinner > 0 ? ((actualRevDinnerPercepatan / targetPercepatanDinner) * 100).toFixed(1) + '%' : '-']);
      sheet.addRow(['TOTAL', 'SEMUA STRATEGI', grandTotalTarget, plannedTargetLunchPenyeimbang + plannedTargetLunchPercepatan + plannedTargetDinnerPenyeimbang + plannedTargetDinnerPercepatan, grandTotalActualMenu, grandTotalActualMenu - grandTotalTarget, grandPct.toFixed(1) + '%']);
      sheet.addRow([]);

      sheet.addRow(['3. DETAIL ITEM MENU STRATEGI TERPILIH']);
      sheet.addRow(['Shift', 'Kategori', 'Nama Menu', 'Harga Satuan (Rp)', 'Target Qty', 'Actual Qty', 'Target Rev (Rp)', 'Actual Rev (Rp)', 'Selisih (Rp)', 'Pencapaian (%)']);

      const allRows = [
        ...lunchPenyeimbangRows.map(r => ({ ...r, shift: 'Lunch', cat: 'Penyeimbang' })),
        ...lunchPercepatanRows.map(r => ({ ...r, shift: 'Lunch', cat: 'Percepatan' })),
        ...dinnerPenyeimbangRows.map(r => ({ ...r, shift: 'Dinner', cat: 'Penyeimbang' })),
        ...dinnerPercepatanRows.map(r => ({ ...r, shift: 'Dinner', cat: 'Percepatan' }))
      ];

      allRows.forEach(r => {
        sheet.addRow([
          r.shift,
          r.cat,
          r.name,
          r.price,
          r.targetQty,
          r.actualQty,
          r.targetRev,
          r.actualRev,
          r.gapRev,
          r.pct.toFixed(1) + '%'
        ]);
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `Menu_Strategi_Kontribusi_${activeOutlet.replace(/\s+/g, '_')}_${selectedDate}.xlsx`);
      showToast?.('Berhasil mengunduh Excel Menu Strategi & Analisa Kontribusi!', 'success');
    } catch (e) {
      console.error(e);
      showToast?.('Gagal mengekspor Excel', 'error');
    }
  };

  // Export to PDF
  const handleExportPDF = () => {
    try {
      const doc = new jsPDF('landscape');
      doc.setFontSize(14);
      doc.text(`JUSTUS STEAKHOUSE - LAPORAN MENU STRATEGI & KONTRIBUSI REVENUE`, 14, 15);
      doc.setFontSize(10);
      doc.text(`Outlet: ${activeOutlet} | Tanggal: ${selectedDate} | COGS Target: Rp ${cogsTargets.strategyTargetRevenue.toLocaleString('id-ID')}`, 14, 22);

      const contribData = [
        ['Lunch', `Rp ${effectiveLunchRevenue.toLocaleString('id-ID')}`, `Rp ${totalActualLunchMenu.toLocaleString('id-ID')}`, `${contribLunchMenuPct.toFixed(1)}%`, `Rp ${actualLunchOutsideMenu.toLocaleString('id-ID')}`, `${contribLunchOutsidePct.toFixed(1)}%`],
        ['Dinner', `Rp ${effectiveDinnerRevenue.toLocaleString('id-ID')}`, `Rp ${totalActualDinnerMenu.toLocaleString('id-ID')}`, `${contribDinnerMenuPct.toFixed(1)}%`, `Rp ${actualDinnerOutsideMenu.toLocaleString('id-ID')}`, `${contribDinnerOutsidePct.toFixed(1)}%`],
        ['Total Day', `Rp ${effectiveTotalRevenue.toLocaleString('id-ID')}`, `Rp ${grandTotalActualMenu.toLocaleString('id-ID')}`, `${contribGrandMenuPct.toFixed(1)}%`, `Rp ${actualGrandOutsideMenu.toLocaleString('id-ID')}`, `${contribGrandOutsidePct.toFixed(1)}%`]
      ];

      autoTable(doc, {
        head: [['Sesi', 'Actual Kasir', 'Menu Strategi', '% Menu', 'Di Luar Menu', '% Luar Menu']],
        body: contribData,
        startY: 28,
        theme: 'grid',
        headStyles: { fillColor: [79, 70, 229] },
        styles: { fontSize: 8 }
      });

      const tableData = [
        ...lunchPenyeimbangRows.map(r => ['Lunch', 'Penyeimbang', r.name, `Rp ${r.price.toLocaleString('id-ID')}`, r.targetQty, r.actualQty, `Rp ${r.targetRev.toLocaleString('id-ID')}`, `Rp ${r.actualRev.toLocaleString('id-ID')}`, `${r.pct.toFixed(1)}%`]),
        ...lunchPercepatanRows.map(r => ['Lunch', 'Percepatan', r.name, `Rp ${r.price.toLocaleString('id-ID')}`, r.targetQty, r.actualQty, `Rp ${r.targetRev.toLocaleString('id-ID')}`, `Rp ${r.actualRev.toLocaleString('id-ID')}`, `${r.pct.toFixed(1)}%`]),
        ...dinnerPenyeimbangRows.map(r => ['Dinner', 'Penyeimbang', r.name, `Rp ${r.price.toLocaleString('id-ID')}`, r.targetQty, r.actualQty, `Rp ${r.targetRev.toLocaleString('id-ID')}`, `Rp ${r.actualRev.toLocaleString('id-ID')}`, `${r.pct.toFixed(1)}%`]),
        ...dinnerPercepatanRows.map(r => ['Dinner', 'Percepatan', r.name, `Rp ${r.price.toLocaleString('id-ID')}`, r.targetQty, r.actualQty, `Rp ${r.targetRev.toLocaleString('id-ID')}`, `Rp ${r.actualRev.toLocaleString('id-ID')}`, `${r.pct.toFixed(1)}%`])
      ];

      // @ts-ignore
      const lastY = (doc as any).lastAutoTable ? (doc as any).lastAutoTable.finalY + 10 : 80;

      autoTable(doc, {
        head: [['Shift', 'Kategori', 'Menu', 'Harga', 'Tgt Qty', 'Act Qty', 'Tgt Rev', 'Act Rev', 'Pencapaian']],
        body: tableData,
        startY: lastY,
        theme: 'grid',
        headStyles: { fillColor: [88, 28, 135] },
        styles: { fontSize: 8 }
      });

      doc.save(`Menu_Strategi_Kontribusi_${activeOutlet.replace(/\s+/g, '_')}_${selectedDate}.pdf`);
      showToast?.('Berhasil mengunduh PDF Menu Strategi!', 'success');
    } catch (e) {
      console.error(e);
      showToast?.('Gagal mengekspor PDF', 'error');
    }
  };

  const openOutletControlStrategi = () => {
    window.dispatchEvent(new CustomEvent('openOutletControl', { detail: { tab: 'STRATEGI' } }));
  };

  return (
    <div className="space-y-6 animate-fade-in pb-20">
      
      {/* HEADER SECTION */}
      <div className="glass-card bg-slate-900 border-white/10 p-6 md:p-8 rounded-[2.5rem] shadow-2xl relative overflow-hidden">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
          <div className="flex items-center gap-4">
            <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-purple-500/20 to-indigo-500/20 border border-purple-500/30 flex items-center justify-center text-purple-400 text-2xl shadow-xl shadow-purple-950/40">
              <i className="fas fa-chess"></i>
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="px-3 py-0.5 rounded-full text-[9px] font-black uppercase tracking-wider bg-purple-500/20 text-purple-300 border border-purple-500/30">
                  REVENUE ACCELERATOR &amp; COGS SYNC
                </span>
                <span className="text-slate-400 text-xs font-bold uppercase">{activeOutlet}</span>
              </div>
              <h2 className="text-2xl md:text-3xl font-black text-white uppercase italic tracking-tighter mt-1">
                FITUR MENU STRATEGI
              </h2>
              <p className="text-xs text-slate-400 mt-1 max-w-xl">
                Target Auto dari COGS, Pengaturan Menu Mendekati Target, dan Analisa Kontribusi terhadap Actual Revenue Lunch &amp; Dinner
              </p>
            </div>
          </div>

          {/* Date Picker & Action Buttons */}
          <div className="flex flex-wrap items-center gap-3 w-full lg:w-auto justify-end">
            <div className="flex items-center gap-2 bg-slate-950 border border-white/10 rounded-2xl px-4 py-2 shadow-inner">
              <i className="fas fa-calendar-day text-purple-400 text-xs"></i>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="bg-transparent text-white font-black text-xs outline-none cursor-pointer"
              />
            </div>

            <button
              onClick={applyCogsTargets}
              className="px-4 py-2.5 rounded-xl bg-purple-600/20 hover:bg-purple-600/30 text-purple-300 border border-purple-500/40 font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm cursor-pointer"
              title="Tarik Target Otomatis dari Fitur COGS"
            >
              <i className="fas fa-sync-alt text-xs"></i>
              <span>Auto dari COGS</span>
            </button>

            <button
              onClick={handleCopyWhatsApp}
              className="px-4 py-2.5 rounded-xl bg-emerald-700/30 hover:bg-emerald-600/40 text-emerald-300 border border-emerald-500/40 font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm cursor-pointer"
              title="Salin Laporan ke Format WhatsApp"
            >
              <i className="fab fa-whatsapp text-xs"></i>
              <span>Salin WA</span>
            </button>

            <button
              onClick={handleExportExcel}
              className="px-3.5 py-2.5 rounded-xl bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/30 font-black text-xs uppercase tracking-wider transition-all flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <i className="fas fa-file-excel text-xs"></i>
              <span>Excel</span>
            </button>

            <button
              onClick={handleExportPDF}
              className="px-3.5 py-2.5 rounded-xl bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 border border-rose-500/30 font-black text-xs uppercase tracking-wider transition-all flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <i className="fas fa-file-pdf text-xs"></i>
              <span>PDF</span>
            </button>

            <button
              onClick={handleSave}
              disabled={isSaving}
              className="px-6 py-2.5 rounded-xl bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-500 hover:to-indigo-500 text-white font-black text-xs uppercase tracking-wider transition-all shadow-xl shadow-purple-950/50 flex items-center gap-2 active:scale-95 cursor-pointer disabled:opacity-50"
            >
              <i className={`fas ${isSaving ? 'fa-spinner fa-spin' : 'fa-save'} text-xs`}></i>
              <span>{isSaving ? 'Menyimpan...' : 'Simpan Data'}</span>
            </button>
          </div>
        </div>

        {/* COGS LIVE SYNC BANNER */}
        <div className="mt-6 pt-4 border-t border-white/10 flex flex-wrap items-center justify-between gap-4 bg-slate-950/60 p-4 rounded-2xl border border-purple-500/20">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-xl bg-emerald-500/20 border border-emerald-500/30 flex items-center justify-center text-emerald-400 text-sm">
              <i className="fas fa-bolt"></i>
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-black uppercase text-emerald-400 tracking-wider">
                  Target Auto dari Fitur COGS
                </span>
                <span className="text-[9px] bg-slate-800 text-slate-300 px-2 py-0.5 rounded-full font-bold">
                  {cogsTargets.percepatanPercent}% Percepatan / {cogsTargets.penyeimbangPercent}% Penyeimbang
                </span>
                <span className="text-[9px] bg-indigo-950/80 text-indigo-300 px-2 py-0.5 rounded-full font-bold">
                  Basis: {cogsTargets.basisName}
                </span>
              </div>
              <p className="text-[11px] text-slate-400 font-bold mt-0.5">
                Target Today Revenue: <strong className="text-white">Rp {cogsTargets.strategyTargetRevenue.toLocaleString('id-ID')}</strong> (Lunch: Rp {cogsTargets.targetLunchNominal.toLocaleString('id-ID')} | Dinner: Rp {cogsTargets.targetDinnerNominal.toLocaleString('id-ID')})
              </p>
            </div>
          </div>
          <button
            onClick={applyCogsTargets}
            className="text-[10px] font-black uppercase px-3 py-1.5 rounded-xl bg-purple-600/30 hover:bg-purple-600/40 text-purple-200 border border-purple-400/30 flex items-center gap-1.5 cursor-pointer transition-all active:scale-95"
          >
            <i className="fas fa-sync-alt"></i> Sinkronisasi Ulang Target COGS
          </button>
        </div>
      </div>

      {/* KPI SUMMARY CARDS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Lunch Card */}
        <div className="glass-card bg-slate-900 border-white/10 p-5 rounded-2xl relative overflow-hidden">
          <div className="flex justify-between items-start">
            <div>
              <span className="text-[10px] font-black uppercase text-amber-400 tracking-wider">Shift Lunch (Menu)</span>
              <h4 className="text-xl font-black text-white mt-1">
                Rp {totalActualLunchMenu.toLocaleString('id-ID')}
              </h4>
            </div>
            <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase border ${pctLunch >= 100 ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30' : pctLunch >= 75 ? 'bg-amber-500/20 text-amber-400 border-amber-500/30' : 'bg-rose-500/20 text-rose-400 border-rose-500/30'}`}>
              {pctLunch.toFixed(1)}%
            </span>
          </div>
          <div className="mt-3 flex justify-between text-[10px] text-slate-400 font-bold border-t border-white/5 pt-2">
            <span>Target: Rp {totalTargetLunch.toLocaleString('id-ID')}</span>
            <span className={totalActualLunchMenu >= totalTargetLunch ? 'text-emerald-400' : 'text-rose-400'}>
              Gap: Rp {(totalActualLunchMenu - totalTargetLunch).toLocaleString('id-ID')}
            </span>
          </div>
        </div>

        {/* Dinner Card */}
        <div className="glass-card bg-slate-900 border-white/10 p-5 rounded-2xl relative overflow-hidden">
          <div className="flex justify-between items-start">
            <div>
              <span className="text-[10px] font-black uppercase text-indigo-400 tracking-wider">Shift Dinner (Menu)</span>
              <h4 className="text-xl font-black text-white mt-1">
                Rp {totalActualDinnerMenu.toLocaleString('id-ID')}
              </h4>
            </div>
            <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase border ${pctDinner >= 100 ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30' : pctDinner >= 75 ? 'bg-amber-500/20 text-amber-400 border-amber-500/30' : 'bg-rose-500/20 text-rose-400 border-rose-500/30'}`}>
              {pctDinner.toFixed(1)}%
            </span>
          </div>
          <div className="mt-3 flex justify-between text-[10px] text-slate-400 font-bold border-t border-white/5 pt-2">
            <span>Target: Rp {totalTargetDinner.toLocaleString('id-ID')}</span>
            <span className={totalActualDinnerMenu >= totalTargetDinner ? 'text-emerald-400' : 'text-rose-400'}>
              Gap: Rp {(totalActualDinnerMenu - totalTargetDinner).toLocaleString('id-ID')}
            </span>
          </div>
        </div>

        {/* Total Penyeimbang */}
        <div className="glass-card bg-slate-900 border-white/10 p-5 rounded-2xl relative overflow-hidden">
          <div className="flex justify-between items-start">
            <div>
              <span className="text-[10px] font-black uppercase text-purple-400 tracking-wider">Menu Penyeimbang</span>
              <h4 className="text-xl font-black text-purple-300 mt-1">
                Rp {totalActualPenyeimbang.toLocaleString('id-ID')}
              </h4>
            </div>
            <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase border ${pctPenyeimbang >= 100 ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30' : 'bg-purple-500/20 text-purple-300 border-purple-500/30'}`}>
              {pctPenyeimbang.toFixed(1)}%
            </span>
          </div>
          <div className="mt-3 flex justify-between text-[10px] text-slate-400 font-bold border-t border-white/5 pt-2">
            <span>Target: Rp {totalTargetPenyeimbang.toLocaleString('id-ID')}</span>
            <span className={totalActualPenyeimbang >= totalTargetPenyeimbang ? 'text-emerald-400' : 'text-slate-400'}>
              Selisih: Rp {(totalActualPenyeimbang - totalTargetPenyeimbang).toLocaleString('id-ID')}
            </span>
          </div>
        </div>

        {/* Total Percepatan */}
        <div className="glass-card bg-slate-900 border-white/10 p-5 rounded-2xl relative overflow-hidden">
          <div className="flex justify-between items-start">
            <div>
              <span className="text-[10px] font-black uppercase text-emerald-400 tracking-wider">Menu Percepatan</span>
              <h4 className="text-xl font-black text-emerald-300 mt-1">
                Rp {totalActualPercepatan.toLocaleString('id-ID')}
              </h4>
            </div>
            <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase border ${pctPercepatan >= 100 ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30' : 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30'}`}>
              {pctPercepatan.toFixed(1)}%
            </span>
          </div>
          <div className="mt-3 flex justify-between text-[10px] text-slate-400 font-bold border-t border-white/5 pt-2">
            <span>Target: Rp {totalTargetPercepatan.toLocaleString('id-ID')}</span>
            <span className={totalActualPercepatan >= totalTargetPercepatan ? 'text-emerald-400' : 'text-slate-400'}>
              Selisih: Rp {(totalActualPercepatan - totalTargetPercepatan).toLocaleString('id-ID')}
            </span>
          </div>
        </div>
      </div>

      {/* SECTION 1: TARGET REVENUE PENYEIMBANG & PERCEPATAN LUNCH & DINNER */}
      <div className="glass-card bg-slate-900 border-white/10 p-6 md:p-8 rounded-[2rem] shadow-xl space-y-6">
        <div className="border-b border-white/10 pb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
          <div>
            <h3 className="text-lg font-black text-white uppercase tracking-tight flex items-center gap-2.5">
              <i className="fas fa-bullseye text-purple-400"></i>
              1. Target Revenue Penyeimbang &amp; Percepatan (Auto dari COGS)
            </h3>
            <p className="text-xs text-slate-400 mt-0.5">
              Target muncul otomatis dari fitur COGS. Anda juga dapat menyesuaikan nominal target secara manual jika diperlukan.
            </p>
          </div>
          <div className="text-right flex items-center gap-4">
            <div>
              <span className="text-[10px] font-bold text-slate-500 uppercase block">Grand Target Strategi:</span>
              <div className="text-lg font-black text-amber-400 font-mono">
                Rp {grandTotalTarget.toLocaleString('id-ID')}
              </div>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* LUNCH TARGETS */}
          <div className="bg-slate-950/60 p-5 rounded-2xl border border-amber-500/20 space-y-4">
            <div className="flex justify-between items-center border-b border-white/5 pb-2">
              <h4 className="text-xs font-black uppercase text-amber-400 tracking-wider flex items-center gap-2">
                <i className="fas fa-sun"></i> Shift Lunch (Siang)
              </h4>
              <span className="text-xs font-black text-white font-mono">
                Total Target: Rp {totalTargetLunch.toLocaleString('id-ID')}
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {/* Lunch Penyeimbang */}
              <div className="bg-slate-900/60 p-3.5 rounded-xl border border-white/5 space-y-2">
                <div className="flex justify-between items-center">
                  <label className="text-[10px] font-black text-purple-300 uppercase tracking-wider flex items-center gap-1.5">
                    <i className="fas fa-scale-balanced text-purple-400 text-xs"></i>
                    Penyeimbang (Lunch)
                  </label>
                  <button
                    type="button"
                    onClick={() => handleAutoSetMenu('LUNCH_PENYEIMBANG')}
                    className="text-[8px] font-black uppercase px-2 py-0.5 rounded bg-purple-500/20 text-purple-300 border border-purple-500/30 hover:bg-purple-500/30 transition-all cursor-pointer"
                    title="Atur kuantitas menu otomatis mendekati nominal target ini"
                  >
                    ⚡ Auto-Set Menu
                  </button>
                </div>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-xs font-bold">Rp</span>
                  <input
                    type="text"
                    value={targetPenyeimbangLunch === 0 ? '' : targetPenyeimbangLunch.toLocaleString('id-ID')}
                    onChange={(e) => {
                      const num = Math.max(0, parseInt(e.target.value.replace(/[^0-9]/g, ''), 10) || 0);
                      setTargetPenyeimbangLunch(num);
                    }}
                    placeholder="0"
                    className="w-full bg-slate-950 border border-white/10 rounded-xl pl-9 pr-3 py-2 text-white font-mono font-bold text-xs outline-none focus:border-amber-400"
                  />
                </div>
                <div className="flex justify-between text-[9px] pt-1 border-t border-white/5 font-mono">
                  <span className="text-slate-400">Rencana Menu: Rp {plannedTargetLunchPenyeimbang.toLocaleString('id-ID')}</span>
                  <span className={Math.abs(plannedTargetLunchPenyeimbang - targetPenyeimbangLunch) <= 15000 ? 'text-emerald-400 font-bold' : 'text-amber-400'}>
                    Selisih: {plannedTargetLunchPenyeimbang >= targetPenyeimbangLunch ? '+' : ''}Rp {(plannedTargetLunchPenyeimbang - targetPenyeimbangLunch).toLocaleString('id-ID')}
                  </span>
                </div>
              </div>

              {/* Lunch Percepatan */}
              <div className="bg-slate-900/60 p-3.5 rounded-xl border border-white/5 space-y-2">
                <div className="flex justify-between items-center">
                  <label className="text-[10px] font-black text-emerald-300 uppercase tracking-wider flex items-center gap-1.5">
                    <i className="fas fa-bolt text-emerald-400 text-xs"></i>
                    Percepatan (Lunch)
                  </label>
                  <button
                    type="button"
                    onClick={() => handleAutoSetMenu('LUNCH_PERCEPATAN')}
                    className="text-[8px] font-black uppercase px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 hover:bg-emerald-500/30 transition-all cursor-pointer"
                    title="Atur kuantitas menu otomatis mendekati nominal target ini"
                  >
                    ⚡ Auto-Set Menu
                  </button>
                </div>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-xs font-bold">Rp</span>
                  <input
                    type="text"
                    value={targetPercepatanLunch === 0 ? '' : targetPercepatanLunch.toLocaleString('id-ID')}
                    onChange={(e) => {
                      const num = Math.max(0, parseInt(e.target.value.replace(/[^0-9]/g, ''), 10) || 0);
                      setTargetPercepatanLunch(num);
                    }}
                    placeholder="0"
                    className="w-full bg-slate-950 border border-white/10 rounded-xl pl-9 pr-3 py-2 text-white font-mono font-bold text-xs outline-none focus:border-amber-400"
                  />
                </div>
                <div className="flex justify-between text-[9px] pt-1 border-t border-white/5 font-mono">
                  <span className="text-slate-400">Rencana Menu: Rp {plannedTargetLunchPercepatan.toLocaleString('id-ID')}</span>
                  <span className={Math.abs(plannedTargetLunchPercepatan - targetPercepatanLunch) <= 15000 ? 'text-emerald-400 font-bold' : 'text-amber-400'}>
                    Selisih: {plannedTargetLunchPercepatan >= targetPercepatanLunch ? '+' : ''}Rp {(plannedTargetLunchPercepatan - targetPercepatanLunch).toLocaleString('id-ID')}
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* DINNER TARGETS */}
          <div className="bg-slate-950/60 p-5 rounded-2xl border border-indigo-500/20 space-y-4">
            <div className="flex justify-between items-center border-b border-white/5 pb-2">
              <h4 className="text-xs font-black uppercase text-indigo-400 tracking-wider flex items-center gap-2">
                <i className="fas fa-moon"></i> Shift Dinner (Malam)
              </h4>
              <span className="text-xs font-black text-white font-mono">
                Total Target: Rp {totalTargetDinner.toLocaleString('id-ID')}
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {/* Dinner Penyeimbang */}
              <div className="bg-slate-900/60 p-3.5 rounded-xl border border-white/5 space-y-2">
                <div className="flex justify-between items-center">
                  <label className="text-[10px] font-black text-purple-300 uppercase tracking-wider flex items-center gap-1.5">
                    <i className="fas fa-scale-balanced text-purple-400 text-xs"></i>
                    Penyeimbang (Dinner)
                  </label>
                  <button
                    type="button"
                    onClick={() => handleAutoSetMenu('DINNER_PENYEIMBANG')}
                    className="text-[8px] font-black uppercase px-2 py-0.5 rounded bg-purple-500/20 text-purple-300 border border-purple-500/30 hover:bg-purple-500/30 transition-all cursor-pointer"
                    title="Atur kuantitas menu otomatis mendekati nominal target ini"
                  >
                    ⚡ Auto-Set Menu
                  </button>
                </div>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-xs font-bold">Rp</span>
                  <input
                    type="text"
                    value={targetPenyeimbangDinner === 0 ? '' : targetPenyeimbangDinner.toLocaleString('id-ID')}
                    onChange={(e) => {
                      const num = Math.max(0, parseInt(e.target.value.replace(/[^0-9]/g, ''), 10) || 0);
                      setTargetPenyeimbangDinner(num);
                    }}
                    placeholder="0"
                    className="w-full bg-slate-900 border border-white/10 rounded-xl pl-9 pr-3 py-2 text-white font-mono font-bold text-xs outline-none focus:border-indigo-400"
                  />
                </div>
                <div className="flex justify-between text-[9px] pt-1 border-t border-white/5 font-mono">
                  <span className="text-slate-400">Rencana Menu: Rp {plannedTargetDinnerPenyeimbang.toLocaleString('id-ID')}</span>
                  <span className={Math.abs(plannedTargetDinnerPenyeimbang - targetPenyeimbangDinner) <= 15000 ? 'text-emerald-400 font-bold' : 'text-amber-400'}>
                    Selisih: {plannedTargetDinnerPenyeimbang >= targetPenyeimbangDinner ? '+' : ''}Rp {(plannedTargetDinnerPenyeimbang - targetPenyeimbangDinner).toLocaleString('id-ID')}
                  </span>
                </div>
              </div>

              {/* Dinner Percepatan */}
              <div className="bg-slate-900/60 p-3.5 rounded-xl border border-white/5 space-y-2">
                <div className="flex justify-between items-center">
                  <label className="text-[10px] font-black text-emerald-300 uppercase tracking-wider flex items-center gap-1.5">
                    <i className="fas fa-bolt text-emerald-400 text-xs"></i>
                    Percepatan (Dinner)
                  </label>
                  <button
                    type="button"
                    onClick={() => handleAutoSetMenu('DINNER_PERCEPATAN')}
                    className="text-[8px] font-black uppercase px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 hover:bg-emerald-500/30 transition-all cursor-pointer"
                    title="Atur kuantitas menu otomatis mendekati nominal target ini"
                  >
                    ⚡ Auto-Set Menu
                  </button>
                </div>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-xs font-bold">Rp</span>
                  <input
                    type="text"
                    value={targetPercepatanDinner === 0 ? '' : targetPercepatanDinner.toLocaleString('id-ID')}
                    onChange={(e) => {
                      const num = Math.max(0, parseInt(e.target.value.replace(/[^0-9]/g, ''), 10) || 0);
                      setTargetPercepatanDinner(num);
                    }}
                    placeholder="0"
                    className="w-full bg-slate-900 border border-white/10 rounded-xl pl-9 pr-3 py-2 text-white font-mono font-bold text-xs outline-none focus:border-indigo-400"
                  />
                </div>
                <div className="flex justify-between text-[9px] pt-1 border-t border-white/5 font-mono">
                  <span className="text-slate-400">Rencana Menu: Rp {plannedTargetDinnerPercepatan.toLocaleString('id-ID')}</span>
                  <span className={Math.abs(plannedTargetDinnerPercepatan - targetPercepatanDinner) <= 15000 ? 'text-emerald-400 font-bold' : 'text-amber-400'}>
                    Selisih: {plannedTargetDinnerPercepatan >= targetPercepatanDinner ? '+' : ''}Rp {(plannedTargetDinnerPercepatan - targetPercepatanDinner).toLocaleString('id-ID')}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* SECTION 2: PILIH & ATUR MENU MENDEKATI TARGET */}
      <div className="glass-card bg-slate-900 border-white/10 p-6 md:p-8 rounded-[2rem] shadow-xl space-y-6">
        <div className="border-b border-white/10 pb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
          <div>
            <h3 className="text-lg font-black text-white uppercase tracking-tight flex items-center gap-2.5">
              <i className="fas fa-utensils text-amber-400"></i>
              2. Pilih &amp; Atur Menu yang Nominalnya Mendekati Target
            </h3>
            <p className="text-xs text-slate-400 mt-0.5">
              Pilih menu dari katalog penyeimbang &amp; percepatan. Gunakan tombol 'Auto-Set Menu' atau atur kuantitas hingga nominal menu mendekati target nominal di atas.
            </p>
          </div>
          <button
            onClick={openOutletControlStrategi}
            className="text-xs text-purple-400 hover:text-purple-300 font-black flex items-center gap-1.5 transition-colors cursor-pointer"
          >
            <i className="fas fa-plus-circle"></i> Tambah Menu Baru di Control Outlet
          </button>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* LUNCH MENU SELECTOR */}
          <div className="bg-slate-950/40 p-5 rounded-2xl border border-white/5 space-y-4">
            <div className="flex justify-between items-center border-b border-white/5 pb-2">
              <h4 className="text-xs font-black uppercase text-amber-400 tracking-wider flex items-center gap-2">
                <i className="fas fa-sun"></i> Menu Pilihan Shift Lunch (Siang)
              </h4>
              <span className="text-[10px] text-slate-400 font-mono">
                Rencana Lunch: <strong className="text-amber-300 font-bold">Rp {(plannedTargetLunchPenyeimbang + plannedTargetLunchPercepatan).toLocaleString('id-ID')}</strong> / Target: Rp {totalTargetLunch.toLocaleString('id-ID')}
              </span>
            </div>

            {/* Lunch Penyeimbang */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <label className="text-[10px] font-black text-purple-300 uppercase tracking-wider flex items-center gap-2">
                  <span>Menu Penyeimbang (Lunch):</span>
                  <span className="text-slate-500 font-bold">{selectedLunchPenyeimbang.length} Terpilih</span>
                </label>
                <div className="flex items-center gap-2">
                  <span className="text-[9px] font-mono text-slate-400">Rp {plannedTargetLunchPenyeimbang.toLocaleString('id-ID')}</span>
                  <button
                    type="button"
                    onClick={() => handleAutoSetMenu('LUNCH_PENYEIMBANG')}
                    className="text-[8px] font-black uppercase px-2 py-0.5 rounded bg-purple-500/20 text-purple-300 border border-purple-500/30 hover:bg-purple-500/40 cursor-pointer"
                  >
                    ⚡ Auto-Set
                  </button>
                </div>
              </div>
              <div className="flex flex-wrap gap-2 max-h-36 overflow-y-auto custom-scrollbar p-1">
                {penyeimbangCatalog.map(item => {
                  const isSelected = selectedLunchPenyeimbang.includes(item.id);
                  const qty = lunchPenyeimbangQtys[item.id]?.targetQty || 0;
                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => toggleSelection(selectedLunchPenyeimbang, setSelectedLunchPenyeimbang, item.id)}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all border flex items-center gap-2 cursor-pointer ${
                        isSelected
                          ? 'bg-purple-600/30 border-purple-500 text-white shadow-sm ring-1 ring-purple-400'
                          : 'bg-slate-900 border-white/10 text-slate-400 hover:text-white'
                      }`}
                    >
                      <i className={`fas ${isSelected ? 'fa-check text-purple-400' : 'fa-plus text-slate-500'} text-[10px]`}></i>
                      <span>{item.name}</span>
                      <span className="text-[9px] text-amber-400 font-mono">Rp {item.price.toLocaleString('id-ID')}</span>
                      {isSelected && qty > 0 && (
                        <span className="text-[9px] bg-purple-500/30 text-purple-200 px-1.5 py-0.2 rounded font-mono font-black">
                          {qty}x
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Lunch Percepatan */}
            <div className="pt-2 border-t border-white/5">
              <div className="flex items-center justify-between mb-2">
                <label className="text-[10px] font-black text-emerald-300 uppercase tracking-wider flex items-center gap-2">
                  <span>Menu Percepatan (Lunch):</span>
                  <span className="text-slate-500 font-bold">{selectedLunchPercepatan.length} Terpilih</span>
                </label>
                <div className="flex items-center gap-2">
                  <span className="text-[9px] font-mono text-slate-400">Rp {plannedTargetLunchPercepatan.toLocaleString('id-ID')}</span>
                  <button
                    type="button"
                    onClick={() => handleAutoSetMenu('LUNCH_PERCEPATAN')}
                    className="text-[8px] font-black uppercase px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 hover:bg-emerald-500/40 cursor-pointer"
                  >
                    ⚡ Auto-Set
                  </button>
                </div>
              </div>
              <div className="flex flex-wrap gap-2 max-h-36 overflow-y-auto custom-scrollbar p-1">
                {percepatanCatalog.map(item => {
                  const isSelected = selectedLunchPercepatan.includes(item.id);
                  const qty = lunchPercepatanQtys[item.id]?.targetQty || 0;
                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => toggleSelection(selectedLunchPercepatan, setSelectedLunchPercepatan, item.id)}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all border flex items-center gap-2 cursor-pointer ${
                        isSelected
                          ? 'bg-emerald-600/30 border-emerald-500 text-white shadow-sm ring-1 ring-emerald-400'
                          : 'bg-slate-900 border-white/10 text-slate-400 hover:text-white'
                      }`}
                    >
                      <i className={`fas ${isSelected ? 'fa-check text-emerald-400' : 'fa-plus text-slate-500'} text-[10px]`}></i>
                      <span>{item.name}</span>
                      <span className="text-[9px] text-amber-400 font-mono">Rp {item.price.toLocaleString('id-ID')}</span>
                      {isSelected && qty > 0 && (
                        <span className="text-[9px] bg-emerald-500/30 text-emerald-200 px-1.5 py-0.2 rounded font-mono font-black">
                          {qty}x
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </div>
          </div>

          {/* DINNER MENU SELECTOR */}
          <div className="bg-slate-950/40 p-5 rounded-2xl border border-white/5 space-y-4">
            <div className="flex justify-between items-center border-b border-white/5 pb-2">
              <h4 className="text-xs font-black uppercase text-indigo-400 tracking-wider flex items-center gap-2">
                <i className="fas fa-moon"></i> Menu Pilihan Shift Dinner (Malam)
              </h4>
              <span className="text-[10px] text-slate-400 font-mono">
                Rencana Dinner: <strong className="text-indigo-300 font-bold">Rp {(plannedTargetDinnerPenyeimbang + plannedTargetDinnerPercepatan).toLocaleString('id-ID')}</strong> / Target: Rp {totalTargetDinner.toLocaleString('id-ID')}
              </span>
            </div>

            {/* Dinner Penyeimbang */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <label className="text-[10px] font-black text-purple-300 uppercase tracking-wider flex items-center gap-2">
                  <span>Menu Penyeimbang (Dinner):</span>
                  <span className="text-slate-500 font-bold">{selectedDinnerPenyeimbang.length} Terpilih</span>
                </label>
                <div className="flex items-center gap-2">
                  <span className="text-[9px] font-mono text-slate-400">Rp {plannedTargetDinnerPenyeimbang.toLocaleString('id-ID')}</span>
                  <button
                    type="button"
                    onClick={() => handleAutoSetMenu('DINNER_PENYEIMBANG')}
                    className="text-[8px] font-black uppercase px-2 py-0.5 rounded bg-purple-500/20 text-purple-300 border border-purple-500/30 hover:bg-purple-500/40 cursor-pointer"
                  >
                    ⚡ Auto-Set
                  </button>
                </div>
              </div>
              <div className="flex flex-wrap gap-2 max-h-36 overflow-y-auto custom-scrollbar p-1">
                {penyeimbangCatalog.map(item => {
                  const isSelected = selectedDinnerPenyeimbang.includes(item.id);
                  const qty = dinnerPenyeimbangQtys[item.id]?.targetQty || 0;
                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => toggleSelection(selectedDinnerPenyeimbang, setSelectedDinnerPenyeimbang, item.id)}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all border flex items-center gap-2 cursor-pointer ${
                        isSelected
                          ? 'bg-purple-600/30 border-purple-500 text-white shadow-sm ring-1 ring-purple-400'
                          : 'bg-slate-900 border-white/10 text-slate-400 hover:text-white'
                      }`}
                    >
                      <i className={`fas ${isSelected ? 'fa-check text-purple-400' : 'fa-plus text-slate-500'} text-[10px]`}></i>
                      <span>{item.name}</span>
                      <span className="text-[9px] text-amber-400 font-mono">Rp {item.price.toLocaleString('id-ID')}</span>
                      {isSelected && qty > 0 && (
                        <span className="text-[9px] bg-purple-500/30 text-purple-200 px-1.5 py-0.2 rounded font-mono font-black">
                          {qty}x
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Dinner Percepatan */}
            <div className="pt-2 border-t border-white/5">
              <div className="flex items-center justify-between mb-2">
                <label className="text-[10px] font-black text-emerald-300 uppercase tracking-wider flex items-center gap-2">
                  <span>Menu Percepatan (Dinner):</span>
                  <span className="text-slate-500 font-bold">{selectedDinnerPercepatan.length} Terpilih</span>
                </label>
                <div className="flex items-center gap-2">
                  <span className="text-[9px] font-mono text-slate-400">Rp {plannedTargetDinnerPercepatan.toLocaleString('id-ID')}</span>
                  <button
                    type="button"
                    onClick={() => handleAutoSetMenu('DINNER_PERCEPATAN')}
                    className="text-[8px] font-black uppercase px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 hover:bg-emerald-500/40 cursor-pointer"
                  >
                    ⚡ Auto-Set
                  </button>
                </div>
              </div>
              <div className="flex flex-wrap gap-2 max-h-36 overflow-y-auto custom-scrollbar p-1">
                {percepatanCatalog.map(item => {
                  const isSelected = selectedDinnerPercepatan.includes(item.id);
                  const qty = dinnerPercepatanQtys[item.id]?.targetQty || 0;
                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => toggleSelection(selectedDinnerPercepatan, setSelectedDinnerPercepatan, item.id)}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all border flex items-center gap-2 cursor-pointer ${
                        isSelected
                          ? 'bg-emerald-600/30 border-emerald-500 text-white shadow-sm ring-1 ring-emerald-400'
                          : 'bg-slate-900 border-white/10 text-slate-400 hover:text-white'
                      }`}
                    >
                      <i className={`fas ${isSelected ? 'fa-check text-emerald-400' : 'fa-plus text-slate-500'} text-[10px]`}></i>
                      <span>{item.name}</span>
                      <span className="text-[9px] text-amber-400 font-mono">Rp {item.price.toLocaleString('id-ID')}</span>
                      {isSelected && qty > 0 && (
                        <span className="text-[9px] bg-emerald-500/30 text-emerald-200 px-1.5 py-0.2 rounded font-mono font-black">
                          {qty}x
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* SECTION 3: TABEL ACTUAL PENCAPAIAN DARI MENU PENYEIMBANG & PERCEPATAN */}
      <div className="glass-card bg-slate-900 border-white/10 p-6 md:p-8 rounded-[2rem] shadow-xl space-y-6">
        <div className="border-b border-white/10 pb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
          <div>
            <h3 className="text-lg font-black text-white uppercase tracking-tight flex items-center gap-2.5">
              <i className="fas fa-table-list text-emerald-400"></i>
              3. Tabel Target Qty &amp; Actual Pencapaian Menu Penyeimbang &amp; Percepatan
            </h3>
            <p className="text-xs text-slate-400 mt-0.5">
              Atur Target Qty agar nominal menu mendekati target, lalu input Actual Qty terjual dari kasir / POS.
            </p>
          </div>
          <div className="text-right">
            <span className="text-[10px] font-bold text-slate-500 uppercase block">Total Penjualan Menu:</span>
            <div className="text-lg font-black text-emerald-400 font-mono">
              Rp {grandTotalActualMenu.toLocaleString('id-ID')}
            </div>
          </div>
        </div>

        <div className="overflow-x-auto rounded-2xl border border-white/10">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-950/80 text-[10px] uppercase font-black tracking-wider text-slate-400 border-b border-white/10">
              <tr>
                <th className="p-3 text-center">Shift</th>
                <th className="p-3 text-center">Kategori</th>
                <th className="p-3">Nama Menu</th>
                <th className="p-3 text-right">Harga Satuan</th>
                <th className="p-3 text-center">Target Qty</th>
                <th className="p-3 text-center">Actual Qty</th>
                <th className="p-3 text-right">Target Rev</th>
                <th className="p-3 text-right">Actual Rev</th>
                <th className="p-3 text-right">Selisih</th>
                <th className="p-3 text-center">% Capai</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-white/5">
              {/* LUNCH PENYEIMBANG ROWS */}
              {lunchPenyeimbangRows.length > 0 && (
                <>
                  <tr className="bg-purple-950/20 text-purple-300 font-black text-[10px] uppercase tracking-wider">
                    <td colSpan={10} className="p-2.5 px-4 flex justify-between items-center">
                      <span>
                        <i className="fas fa-sun mr-2 text-amber-400"></i>
                        LUNCH - MENU PENYEIMBANG ({lunchPenyeimbangRows.length} Menu)
                      </span>
                      <span className="font-mono text-purple-200">
                        Target Segmen: Rp {targetPenyeimbangLunch.toLocaleString('id-ID')} | Nominal Rencana: Rp {plannedTargetLunchPenyeimbang.toLocaleString('id-ID')}
                      </span>
                    </td>
                  </tr>
                  {lunchPenyeimbangRows.map(row => (
                    <tr key={`l_pen_${row.id}`} className="hover:bg-white/[0.02] transition-colors">
                      <td className="p-3 text-center text-amber-400 font-bold text-[10px]">Lunch</td>
                      <td className="p-3 text-center">
                        <span className="px-2 py-0.5 rounded-full text-[8px] font-black bg-purple-500/20 text-purple-300 border border-purple-500/30 uppercase">
                          Penyeimbang
                        </span>
                      </td>
                      <td className="p-3 font-bold text-white">{row.name}</td>
                      <td className="p-3 text-right font-mono text-slate-300">Rp {row.price.toLocaleString('id-ID')}</td>
                      <td className="p-2 text-center">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              const val = Math.max(0, row.targetQty - 1);
                              setLunchPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >-</button>
                          <input
                            type="number"
                            min="0"
                            value={row.targetQty === 0 ? '' : row.targetQty}
                            onChange={(e) => {
                              const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                              setLunchPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            placeholder="0"
                            className="w-16 bg-slate-950 border border-white/10 rounded-lg p-1.5 text-center text-white font-mono font-bold text-xs outline-none focus:border-purple-400"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const val = row.targetQty + 1;
                              setLunchPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >+</button>
                        </div>
                      </td>
                      <td className="p-2 text-center">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              const val = Math.max(0, row.actualQty - 1);
                              setLunchPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >-</button>
                          <input
                            type="number"
                            min="0"
                            value={row.actualQty === 0 ? '' : row.actualQty}
                            onChange={(e) => {
                              const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                              setLunchPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            placeholder="0"
                            className="w-16 bg-slate-950 border border-purple-500/30 rounded-lg p-1.5 text-center text-emerald-400 font-mono font-bold text-xs outline-none focus:border-emerald-400"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const val = row.actualQty + 1;
                              setLunchPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >+</button>
                        </div>
                      </td>
                      <td className="p-3 text-right font-mono text-slate-400">Rp {row.targetRev.toLocaleString('id-ID')}</td>
                      <td className="p-3 text-right font-mono font-black text-emerald-400">Rp {row.actualRev.toLocaleString('id-ID')}</td>
                      <td className={`p-3 text-right font-mono font-bold ${row.gapRev >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                        {row.gapRev >= 0 ? '+' : ''}Rp {row.gapRev.toLocaleString('id-ID')}
                      </td>
                      <td className="p-3 text-center">
                        <span className={`px-2 py-0.5 rounded-full text-[9px] font-black ${row.pct >= 100 ? 'bg-emerald-500/20 text-emerald-400' : row.pct >= 75 ? 'bg-amber-500/20 text-amber-400' : 'bg-rose-500/20 text-rose-400'}`}>
                          {row.pct.toFixed(0)}%
                        </span>
                      </td>
                    </tr>
                  ))}
                </>
              )}

              {/* LUNCH PERCEPATAN ROWS */}
              {lunchPercepatanRows.length > 0 && (
                <>
                  <tr className="bg-emerald-950/20 text-emerald-300 font-black text-[10px] uppercase tracking-wider">
                    <td colSpan={10} className="p-2.5 px-4 flex justify-between items-center">
                      <span>
                        <i className="fas fa-sun mr-2 text-amber-400"></i>
                        LUNCH - MENU PERCEPATAN ({lunchPercepatanRows.length} Menu)
                      </span>
                      <span className="font-mono text-emerald-200">
                        Target Segmen: Rp {targetPercepatanLunch.toLocaleString('id-ID')} | Nominal Rencana: Rp {plannedTargetLunchPercepatan.toLocaleString('id-ID')}
                      </span>
                    </td>
                  </tr>
                  {lunchPercepatanRows.map(row => (
                    <tr key={`l_per_${row.id}`} className="hover:bg-white/[0.02] transition-colors">
                      <td className="p-3 text-center text-amber-400 font-bold text-[10px]">Lunch</td>
                      <td className="p-3 text-center">
                        <span className="px-2 py-0.5 rounded-full text-[8px] font-black bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 uppercase">
                          Percepatan
                        </span>
                      </td>
                      <td className="p-3 font-bold text-white">{row.name}</td>
                      <td className="p-3 text-right font-mono text-slate-300">Rp {row.price.toLocaleString('id-ID')}</td>
                      <td className="p-2 text-center">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              const val = Math.max(0, row.targetQty - 1);
                              setLunchPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >-</button>
                          <input
                            type="number"
                            min="0"
                            value={row.targetQty === 0 ? '' : row.targetQty}
                            onChange={(e) => {
                              const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                              setLunchPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            placeholder="0"
                            className="w-16 bg-slate-950 border border-white/10 rounded-lg p-1.5 text-center text-white font-mono font-bold text-xs outline-none focus:border-emerald-400"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const val = row.targetQty + 1;
                              setLunchPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >+</button>
                        </div>
                      </td>
                      <td className="p-2 text-center">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              const val = Math.max(0, row.actualQty - 1);
                              setLunchPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >-</button>
                          <input
                            type="number"
                            min="0"
                            value={row.actualQty === 0 ? '' : row.actualQty}
                            onChange={(e) => {
                              const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                              setLunchPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            placeholder="0"
                            className="w-16 bg-slate-950 border border-emerald-500/30 rounded-lg p-1.5 text-center text-emerald-400 font-mono font-bold text-xs outline-none focus:border-emerald-400"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const val = row.actualQty + 1;
                              setLunchPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >+</button>
                        </div>
                      </td>
                      <td className="p-3 text-right font-mono text-slate-400">Rp {row.targetRev.toLocaleString('id-ID')}</td>
                      <td className="p-3 text-right font-mono font-black text-emerald-400">Rp {row.actualRev.toLocaleString('id-ID')}</td>
                      <td className={`p-3 text-right font-mono font-bold ${row.gapRev >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                        {row.gapRev >= 0 ? '+' : ''}Rp {row.gapRev.toLocaleString('id-ID')}
                      </td>
                      <td className="p-3 text-center">
                        <span className={`px-2 py-0.5 rounded-full text-[9px] font-black ${row.pct >= 100 ? 'bg-emerald-500/20 text-emerald-400' : row.pct >= 75 ? 'bg-amber-500/20 text-amber-400' : 'bg-rose-500/20 text-rose-400'}`}>
                          {row.pct.toFixed(0)}%
                        </span>
                      </td>
                    </tr>
                  ))}
                </>
              )}

              {/* Subtotal Lunch */}
              <tr className="bg-amber-500/10 font-black text-xs text-amber-300 border-t border-b border-amber-500/20">
                <td colSpan={6} className="p-3 px-4 uppercase tracking-wider">
                  SUBTOTAL ACTUAL PENJUALAN MENU LUNCH (SIANG)
                </td>
                <td className="p-3 text-right font-mono">Rp {totalTargetLunch.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-amber-400">Rp {totalActualLunchMenu.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono">
                  {totalActualLunchMenu >= totalTargetLunch ? '+' : ''}Rp {(totalActualLunchMenu - totalTargetLunch).toLocaleString('id-ID')}
                </td>
                <td className="p-3 text-center">{pctLunch.toFixed(1)}%</td>
              </tr>

              {/* DINNER PENYEIMBANG ROWS */}
              {dinnerPenyeimbangRows.length > 0 && (
                <>
                  <tr className="bg-purple-950/20 text-purple-300 font-black text-[10px] uppercase tracking-wider">
                    <td colSpan={10} className="p-2.5 px-4 flex justify-between items-center">
                      <span>
                        <i className="fas fa-moon mr-2 text-indigo-400"></i>
                        DINNER - MENU PENYEIMBANG ({dinnerPenyeimbangRows.length} Menu)
                      </span>
                      <span className="font-mono text-purple-200">
                        Target Segmen: Rp {targetPenyeimbangDinner.toLocaleString('id-ID')} | Nominal Rencana: Rp {plannedTargetDinnerPenyeimbang.toLocaleString('id-ID')}
                      </span>
                    </td>
                  </tr>
                  {dinnerPenyeimbangRows.map(row => (
                    <tr key={`d_pen_${row.id}`} className="hover:bg-white/[0.02] transition-colors">
                      <td className="p-3 text-center text-indigo-400 font-bold text-[10px]">Dinner</td>
                      <td className="p-3 text-center">
                        <span className="px-2 py-0.5 rounded-full text-[8px] font-black bg-purple-500/20 text-purple-300 border border-purple-500/30 uppercase">
                          Penyeimbang
                        </span>
                      </td>
                      <td className="p-3 font-bold text-white">{row.name}</td>
                      <td className="p-3 text-right font-mono text-slate-300">Rp {row.price.toLocaleString('id-ID')}</td>
                      <td className="p-2 text-center">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              const val = Math.max(0, row.targetQty - 1);
                              setDinnerPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >-</button>
                          <input
                            type="number"
                            min="0"
                            value={row.targetQty === 0 ? '' : row.targetQty}
                            onChange={(e) => {
                              const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                              setDinnerPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            placeholder="0"
                            className="w-16 bg-slate-950 border border-white/10 rounded-lg p-1.5 text-center text-white font-mono font-bold text-xs outline-none focus:border-indigo-400"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const val = row.targetQty + 1;
                              setDinnerPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >+</button>
                        </div>
                      </td>
                      <td className="p-2 text-center">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              const val = Math.max(0, row.actualQty - 1);
                              setDinnerPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >-</button>
                          <input
                            type="number"
                            min="0"
                            value={row.actualQty === 0 ? '' : row.actualQty}
                            onChange={(e) => {
                              const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                              setDinnerPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            placeholder="0"
                            className="w-16 bg-slate-950 border border-purple-500/30 rounded-lg p-1.5 text-center text-emerald-400 font-mono font-bold text-xs outline-none focus:border-emerald-400"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const val = row.actualQty + 1;
                              setDinnerPenyeimbangQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >+</button>
                        </div>
                      </td>
                      <td className="p-3 text-right font-mono text-slate-400">Rp {row.targetRev.toLocaleString('id-ID')}</td>
                      <td className="p-3 text-right font-mono font-black text-emerald-400">Rp {row.actualRev.toLocaleString('id-ID')}</td>
                      <td className={`p-3 text-right font-mono font-bold ${row.gapRev >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                        {row.gapRev >= 0 ? '+' : ''}Rp {row.gapRev.toLocaleString('id-ID')}
                      </td>
                      <td className="p-3 text-center">
                        <span className={`px-2 py-0.5 rounded-full text-[9px] font-black ${row.pct >= 100 ? 'bg-emerald-500/20 text-emerald-400' : row.pct >= 75 ? 'bg-amber-500/20 text-amber-400' : 'bg-rose-500/20 text-rose-400'}`}>
                          {row.pct.toFixed(0)}%
                        </span>
                      </td>
                    </tr>
                  ))}
                </>
              )}

              {/* DINNER PERCEPATAN ROWS */}
              {dinnerPercepatanRows.length > 0 && (
                <>
                  <tr className="bg-emerald-950/20 text-emerald-300 font-black text-[10px] uppercase tracking-wider">
                    <td colSpan={10} className="p-2.5 px-4 flex justify-between items-center">
                      <span>
                        <i className="fas fa-moon mr-2 text-indigo-400"></i>
                        DINNER - MENU PERCEPATAN ({dinnerPercepatanRows.length} Menu)
                      </span>
                      <span className="font-mono text-emerald-200">
                        Target Segmen: Rp {targetPercepatanDinner.toLocaleString('id-ID')} | Nominal Rencana: Rp {plannedTargetDinnerPercepatan.toLocaleString('id-ID')}
                      </span>
                    </td>
                  </tr>
                  {dinnerPercepatanRows.map(row => (
                    <tr key={`d_per_${row.id}`} className="hover:bg-white/[0.02] transition-colors">
                      <td className="p-3 text-center text-indigo-400 font-bold text-[10px]">Dinner</td>
                      <td className="p-3 text-center">
                        <span className="px-2 py-0.5 rounded-full text-[8px] font-black bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 uppercase">
                          Percepatan
                        </span>
                      </td>
                      <td className="p-3 font-bold text-white">{row.name}</td>
                      <td className="p-3 text-right font-mono text-slate-300">Rp {row.price.toLocaleString('id-ID')}</td>
                      <td className="p-2 text-center">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              const val = Math.max(0, row.targetQty - 1);
                              setDinnerPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >-</button>
                          <input
                            type="number"
                            min="0"
                            value={row.targetQty === 0 ? '' : row.targetQty}
                            onChange={(e) => {
                              const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                              setDinnerPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            placeholder="0"
                            className="w-16 bg-slate-950 border border-white/10 rounded-lg p-1.5 text-center text-white font-mono font-bold text-xs outline-none focus:border-indigo-400"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const val = row.targetQty + 1;
                              setDinnerPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { actualQty: 0 }), targetQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >+</button>
                        </div>
                      </td>
                      <td className="p-2 text-center">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => {
                              const val = Math.max(0, row.actualQty - 1);
                              setDinnerPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >-</button>
                          <input
                            type="number"
                            min="0"
                            value={row.actualQty === 0 ? '' : row.actualQty}
                            onChange={(e) => {
                              const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                              setDinnerPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            placeholder="0"
                            className="w-16 bg-slate-950 border border-emerald-500/30 rounded-lg p-1.5 text-center text-emerald-400 font-mono font-bold text-xs outline-none focus:border-emerald-400"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const val = row.actualQty + 1;
                              setDinnerPercepatanQtys(prev => ({
                                ...prev,
                                [row.id]: { ...(prev[row.id] || { targetQty: 0 }), actualQty: val }
                              }));
                            }}
                            className="w-5 h-5 rounded bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center font-bold text-[10px]"
                          >+</button>
                        </div>
                      </td>
                      <td className="p-3 text-right font-mono text-slate-400">Rp {row.targetRev.toLocaleString('id-ID')}</td>
                      <td className="p-3 text-right font-mono font-black text-emerald-400">Rp {row.actualRev.toLocaleString('id-ID')}</td>
                      <td className={`p-3 text-right font-mono font-bold ${row.gapRev >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                        {row.gapRev >= 0 ? '+' : ''}Rp {row.gapRev.toLocaleString('id-ID')}
                      </td>
                      <td className="p-3 text-center">
                        <span className={`px-2 py-0.5 rounded-full text-[9px] font-black ${row.pct >= 100 ? 'bg-emerald-500/20 text-emerald-400' : row.pct >= 75 ? 'bg-amber-500/20 text-amber-400' : 'bg-rose-500/20 text-rose-400'}`}>
                          {row.pct.toFixed(0)}%
                        </span>
                      </td>
                    </tr>
                  ))}
                </>
              )}

              {/* Subtotal Dinner */}
              <tr className="bg-indigo-500/10 font-black text-xs text-indigo-300 border-t border-b border-indigo-500/20">
                <td colSpan={6} className="p-3 px-4 uppercase tracking-wider">
                  SUBTOTAL ACTUAL PENJUALAN MENU DINNER (MALAM)
                </td>
                <td className="p-3 text-right font-mono">Rp {totalTargetDinner.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-indigo-300">Rp {totalActualDinnerMenu.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono">
                  {totalActualDinnerMenu >= totalTargetDinner ? '+' : ''}Rp {(totalActualDinnerMenu - totalTargetDinner).toLocaleString('id-ID')}
                </td>
                <td className="p-3 text-center">{pctDinner.toFixed(1)}%</td>
              </tr>

              {/* GRAND TOTAL ROW */}
              <tr className="bg-slate-950 font-black text-sm text-white border-t-2 border-purple-500/50">
                <td colSpan={6} className="p-4 px-4 uppercase tracking-wider text-purple-300">
                  <i className="fas fa-crown text-amber-400 mr-2"></i>
                  GRAND TOTAL PENJUALAN MENU STRATEGI (LUNCH + DINNER)
                </td>
                <td className="p-4 text-right font-mono text-amber-400">Rp {grandTotalTarget.toLocaleString('id-ID')}</td>
                <td className="p-4 text-right font-mono text-emerald-400 text-base">Rp {grandTotalActualMenu.toLocaleString('id-ID')}</td>
                <td className={`p-4 text-right font-mono text-base ${grandTotalActualMenu >= grandTotalTarget ? 'text-emerald-400' : 'text-rose-400'}`}>
                  {grandTotalActualMenu >= grandTotalTarget ? '+' : ''}Rp {(grandTotalActualMenu - grandTotalTarget).toLocaleString('id-ID')}
                </td>
                <td className="p-4 text-center">
                  <span className={`px-3 py-1 rounded-full text-xs font-black ${grandPct >= 100 ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/40' : 'bg-rose-500/20 text-rose-400 border border-rose-500/40'}`}>
                    {grandPct.toFixed(1)}%
                  </span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      {/* SECTION 4: ANALISA KONTRIBUSI TERHADAP ACTUAL REVENUE LUNCH & DINNER DI LUAR PENJUALAN MENU */}
      <div className="glass-card bg-slate-900 border-indigo-500/30 p-6 md:p-8 rounded-[2rem] shadow-2xl space-y-6">
        <div className="border-b border-white/10 pb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="px-3 py-0.5 rounded-full text-[9px] font-black uppercase tracking-wider bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                ANALISA REVENUE UTAMA
              </span>
              <span className="text-slate-400 text-xs font-bold">POS / Kasir vs Menu Strategi</span>
            </div>
            <h3 className="text-lg md:text-xl font-black text-white uppercase tracking-tight flex items-center gap-2.5 mt-1">
              <i className="fas fa-chart-pie text-indigo-400"></i>
              4. Analisa Kontribusi Actual Revenue Lunch &amp; Dinner (Penjualan Menu vs Di Luar Penjualan Menu)
            </h3>
            <p className="text-xs text-slate-400 mt-0.5">
              Perbandingan seberapa besar kontribusi pencapaian menu strategi terhadap omset kasir dan berapa porsi revenue di luar penjualan menu.
            </p>
          </div>

          <div className="flex items-center gap-2 bg-slate-950 p-2 rounded-xl border border-white/10">
            <label className="text-[10px] font-bold text-slate-400 flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={isManualRevenueOverride}
                onChange={(e) => setIsManualRevenueOverride(e.target.checked)}
                className="rounded accent-purple-500"
              />
              <span>Input Manual Revenue</span>
            </label>
          </div>
        </div>

        {/* Input Revenue Kasir If Overriding */}
        {isManualRevenueOverride && (
          <div className="bg-slate-950/80 p-4 rounded-xl border border-indigo-500/30 grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <label className="block text-[10px] font-black text-amber-400 uppercase tracking-wider mb-1">
                Actual Revenue Lunch (Kasir / POS):
              </label>
              <div className="relative">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-xs font-bold">Rp</span>
                <input
                  type="text"
                  value={actualLunchRevenueInput === 0 ? '' : actualLunchRevenueInput.toLocaleString('id-ID')}
                  onChange={(e) => {
                    const num = Math.max(0, parseInt(e.target.value.replace(/[^0-9]/g, ''), 10) || 0);
                    setActualLunchRevenueInput(num);
                  }}
                  placeholder="0"
                  className="w-full bg-slate-900 border border-white/10 rounded-xl pl-9 pr-3 py-2 text-white font-mono font-bold text-xs outline-none focus:border-amber-400"
                />
              </div>
            </div>
            <div>
              <label className="block text-[10px] font-black text-indigo-400 uppercase tracking-wider mb-1">
                Actual Revenue Dinner (Kasir / POS):
              </label>
              <div className="relative">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-xs font-bold">Rp</span>
                <input
                  type="text"
                  value={actualDinnerRevenueInput === 0 ? '' : actualDinnerRevenueInput.toLocaleString('id-ID')}
                  onChange={(e) => {
                    const num = Math.max(0, parseInt(e.target.value.replace(/[^0-9]/g, ''), 10) || 0);
                    setActualDinnerRevenueInput(num);
                  }}
                  placeholder="0"
                  className="w-full bg-slate-900 border border-white/10 rounded-xl pl-9 pr-3 py-2 text-white font-mono font-bold text-xs outline-none focus:border-indigo-400"
                />
              </div>
            </div>
          </div>
        )}

        {/* 3 COMPARISON CARDS */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          {/* LUNCH CARD */}
          <div className="bg-slate-950/70 p-5 rounded-2xl border border-amber-500/20 space-y-4">
            <div className="flex justify-between items-center border-b border-white/5 pb-2">
              <h4 className="text-xs font-black uppercase text-amber-400 tracking-wider flex items-center gap-2">
                <i className="fas fa-sun"></i> Shift Lunch (Siang)
              </h4>
              <span className="text-[10px] font-black text-slate-400 uppercase">
                Kasir: Rp {effectiveLunchRevenue.toLocaleString('id-ID')}
              </span>
            </div>

            {/* Dual visual progress bar */}
            <div className="space-y-1.5">
              <div className="h-3 w-full bg-slate-900 rounded-full overflow-hidden flex">
                <div
                  style={{ width: `${Math.min(100, contribLunchMenuPct)}%` }}
                  className="bg-gradient-to-r from-amber-500 to-purple-500 h-full transition-all"
                  title={`Menu Strategi: ${contribLunchMenuPct.toFixed(1)}%`}
                ></div>
                <div
                  style={{ width: `${Math.min(100, contribLunchOutsidePct)}%` }}
                  className="bg-slate-700 h-full transition-all"
                  title={`Di Luar Menu: ${contribLunchOutsidePct.toFixed(1)}%`}
                ></div>
              </div>
              <div className="flex justify-between text-[9px] text-slate-400 font-bold">
                <span className="text-amber-300">Menu: {contribLunchMenuPct.toFixed(1)}%</span>
                <span className="text-slate-400">Luar Menu: {contribLunchOutsidePct.toFixed(1)}%</span>
              </div>
            </div>

            <div className="space-y-2.5 pt-2 text-xs">
              <div className="flex justify-between items-center p-2.5 rounded-xl bg-slate-900/60 border border-white/5">
                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-amber-400"></div>
                  <span className="text-[11px] font-bold text-slate-300">Penjualan Menu Strategi</span>
                </div>
                <div className="text-right">
                  <div className="font-mono font-bold text-white">Rp {totalActualLunchMenu.toLocaleString('id-ID')}</div>
                  <div className="text-[9px] font-black text-amber-400">{contribLunchMenuPct.toFixed(1)}% Kontribusi</div>
                </div>
              </div>

              <div className="flex justify-between items-center p-2.5 rounded-xl bg-slate-900/60 border border-white/5">
                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-slate-500"></div>
                  <span className="text-[11px] font-bold text-slate-300">Di Luar Penjualan Menu</span>
                </div>
                <div className="text-right">
                  <div className="font-mono font-bold text-slate-300">Rp {actualLunchOutsideMenu.toLocaleString('id-ID')}</div>
                  <div className="text-[9px] font-black text-slate-400">{contribLunchOutsidePct.toFixed(1)}% Kontribusi</div>
                </div>
              </div>
            </div>
          </div>

          {/* DINNER CARD */}
          <div className="bg-slate-950/70 p-5 rounded-2xl border border-indigo-500/20 space-y-4">
            <div className="flex justify-between items-center border-b border-white/5 pb-2">
              <h4 className="text-xs font-black uppercase text-indigo-400 tracking-wider flex items-center gap-2">
                <i className="fas fa-moon"></i> Shift Dinner (Malam)
              </h4>
              <span className="text-[10px] font-black text-slate-400 uppercase">
                Kasir: Rp {effectiveDinnerRevenue.toLocaleString('id-ID')}
              </span>
            </div>

            {/* Dual visual progress bar */}
            <div className="space-y-1.5">
              <div className="h-3 w-full bg-slate-900 rounded-full overflow-hidden flex">
                <div
                  style={{ width: `${Math.min(100, contribDinnerMenuPct)}%` }}
                  className="bg-gradient-to-r from-indigo-500 to-purple-500 h-full transition-all"
                  title={`Menu Strategi: ${contribDinnerMenuPct.toFixed(1)}%`}
                ></div>
                <div
                  style={{ width: `${Math.min(100, contribDinnerOutsidePct)}%` }}
                  className="bg-slate-700 h-full transition-all"
                  title={`Di Luar Menu: ${contribDinnerOutsidePct.toFixed(1)}%`}
                ></div>
              </div>
              <div className="flex justify-between text-[9px] text-slate-400 font-bold">
                <span className="text-indigo-300">Menu: {contribDinnerMenuPct.toFixed(1)}%</span>
                <span className="text-slate-400">Luar Menu: {contribDinnerOutsidePct.toFixed(1)}%</span>
              </div>
            </div>

            <div className="space-y-2.5 pt-2 text-xs">
              <div className="flex justify-between items-center p-2.5 rounded-xl bg-slate-900/60 border border-white/5">
                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-indigo-400"></div>
                  <span className="text-[11px] font-bold text-slate-300">Penjualan Menu Strategi</span>
                </div>
                <div className="text-right">
                  <div className="font-mono font-bold text-white">Rp {totalActualDinnerMenu.toLocaleString('id-ID')}</div>
                  <div className="text-[9px] font-black text-indigo-400">{contribDinnerMenuPct.toFixed(1)}% Kontribusi</div>
                </div>
              </div>

              <div className="flex justify-between items-center p-2.5 rounded-xl bg-slate-900/60 border border-white/5">
                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-slate-500"></div>
                  <span className="text-[11px] font-bold text-slate-300">Di Luar Penjualan Menu</span>
                </div>
                <div className="text-right">
                  <div className="font-mono font-bold text-slate-300">Rp {actualDinnerOutsideMenu.toLocaleString('id-ID')}</div>
                  <div className="text-[9px] font-black text-slate-400">{contribDinnerOutsidePct.toFixed(1)}% Kontribusi</div>
                </div>
              </div>
            </div>
          </div>

          {/* TOTAL DAY CARD */}
          <div className="bg-slate-950/70 p-5 rounded-2xl border border-purple-500/30 space-y-4 shadow-lg shadow-purple-950/20">
            <div className="flex justify-between items-center border-b border-white/5 pb-2">
              <h4 className="text-xs font-black uppercase text-purple-300 tracking-wider flex items-center gap-2">
                <i className="fas fa-crown text-amber-400"></i> Total Harian (Full Day)
              </h4>
              <span className="text-[10px] font-black text-white font-mono">
                Rp {effectiveTotalRevenue.toLocaleString('id-ID')}
              </span>
            </div>

            {/* Dual visual progress bar */}
            <div className="space-y-1.5">
              <div className="h-3 w-full bg-slate-900 rounded-full overflow-hidden flex">
                <div
                  style={{ width: `${Math.min(100, contribGrandMenuPct)}%` }}
                  className="bg-gradient-to-r from-emerald-500 to-indigo-500 h-full transition-all"
                  title={`Menu Strategi: ${contribGrandMenuPct.toFixed(1)}%`}
                ></div>
                <div
                  style={{ width: `${Math.min(100, contribGrandOutsidePct)}%` }}
                  className="bg-slate-700 h-full transition-all"
                  title={`Di Luar Menu: ${contribGrandOutsidePct.toFixed(1)}%`}
                ></div>
              </div>
              <div className="flex justify-between text-[9px] text-slate-400 font-bold">
                <span className="text-emerald-400 font-bold">Menu: {contribGrandMenuPct.toFixed(1)}%</span>
                <span className="text-slate-400">Luar Menu: {contribGrandOutsidePct.toFixed(1)}%</span>
              </div>
            </div>

            <div className="space-y-2.5 pt-2 text-xs">
              <div className="flex justify-between items-center p-2.5 rounded-xl bg-purple-950/30 border border-purple-500/20">
                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-emerald-400"></div>
                  <span className="text-[11px] font-bold text-white">Grand Menu Strategi</span>
                </div>
                <div className="text-right">
                  <div className="font-mono font-bold text-emerald-400">Rp {grandTotalActualMenu.toLocaleString('id-ID')}</div>
                  <div className="text-[9px] font-black text-emerald-300">{contribGrandMenuPct.toFixed(1)}% Kontribusi</div>
                </div>
              </div>

              <div className="flex justify-between items-center p-2.5 rounded-xl bg-slate-900/60 border border-white/5">
                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-slate-400"></div>
                  <span className="text-[11px] font-bold text-slate-300">Total Di Luar Menu</span>
                </div>
                <div className="text-right">
                  <div className="font-mono font-bold text-slate-300">Rp {actualGrandOutsideMenu.toLocaleString('id-ID')}</div>
                  <div className="text-[9px] font-black text-slate-400">{contribGrandOutsidePct.toFixed(1)}% Kontribusi</div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* COMPARISON MATRIX TABLE */}
        <div className="overflow-x-auto rounded-2xl border border-white/10 mt-4">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-950/90 text-[10px] uppercase font-black tracking-wider text-slate-400 border-b border-white/10">
              <tr>
                <th className="p-3">Sesi / Shift</th>
                <th className="p-3 text-right">Actual Kasir (POS)</th>
                <th className="p-3 text-right">Penyeimbang (Act)</th>
                <th className="p-3 text-right">Percepatan (Act)</th>
                <th className="p-3 text-right">Total Menu Strategi</th>
                <th className="p-3 text-center">% Kontribusi Menu</th>
                <th className="p-3 text-right">Di Luar Menu Strategi</th>
                <th className="p-3 text-center">% Kontribusi Luar Menu</th>
                <th className="p-3 text-center">Status Up-Selling</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-white/5 text-xs">
              {/* Lunch Row */}
              <tr className="hover:bg-white/[0.02]">
                <td className="p-3 font-bold text-amber-400 flex items-center gap-2">
                  <i className="fas fa-sun text-xs"></i> Shift Lunch
                </td>
                <td className="p-3 text-right font-mono font-bold text-white">Rp {effectiveLunchRevenue.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-purple-300">Rp {actualRevLunchPenyeimbang.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-emerald-300">Rp {actualRevLunchPercepatan.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono font-black text-amber-300">Rp {totalActualLunchMenu.toLocaleString('id-ID')}</td>
                <td className="p-3 text-center">
                  <span className={`px-2 py-0.5 rounded-full text-[9px] font-black ${contribLunchMenuPct >= 20 ? 'bg-emerald-500/20 text-emerald-400' : 'bg-amber-500/20 text-amber-400'}`}>
                    {contribLunchMenuPct.toFixed(1)}%
                  </span>
                </td>
                <td className="p-3 text-right font-mono text-slate-400">Rp {actualLunchOutsideMenu.toLocaleString('id-ID')}</td>
                <td className="p-3 text-center text-slate-400 font-mono">{contribLunchOutsidePct.toFixed(1)}%</td>
                <td className="p-3 text-center">
                  <span className={`px-2 py-0.5 rounded-full text-[8px] font-black uppercase ${totalActualLunchMenu >= totalTargetLunch ? 'bg-emerald-500/20 text-emerald-300' : 'bg-slate-800 text-slate-400'}`}>
                    {totalActualLunchMenu >= totalTargetLunch ? 'Target Tercapai' : 'Optimalisasi Diperlukan'}
                  </span>
                </td>
              </tr>

              {/* Dinner Row */}
              <tr className="hover:bg-white/[0.02]">
                <td className="p-3 font-bold text-indigo-400 flex items-center gap-2">
                  <i className="fas fa-moon text-xs"></i> Shift Dinner
                </td>
                <td className="p-3 text-right font-mono font-bold text-white">Rp {effectiveDinnerRevenue.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-purple-300">Rp {actualRevDinnerPenyeimbang.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-emerald-300">Rp {actualRevDinnerPercepatan.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono font-black text-indigo-300">Rp {totalActualDinnerMenu.toLocaleString('id-ID')}</td>
                <td className="p-3 text-center">
                  <span className={`px-2 py-0.5 rounded-full text-[9px] font-black ${contribDinnerMenuPct >= 25 ? 'bg-emerald-500/20 text-emerald-400' : 'bg-indigo-500/20 text-indigo-300'}`}>
                    {contribDinnerMenuPct.toFixed(1)}%
                  </span>
                </td>
                <td className="p-3 text-right font-mono text-slate-400">Rp {actualDinnerOutsideMenu.toLocaleString('id-ID')}</td>
                <td className="p-3 text-center text-slate-400 font-mono">{contribDinnerOutsidePct.toFixed(1)}%</td>
                <td className="p-3 text-center">
                  <span className={`px-2 py-0.5 rounded-full text-[8px] font-black uppercase ${totalActualDinnerMenu >= totalTargetDinner ? 'bg-emerald-500/20 text-emerald-300' : 'bg-slate-800 text-slate-400'}`}>
                    {totalActualDinnerMenu >= totalTargetDinner ? 'Target Tercapai' : 'Optimalisasi Diperlukan'}
                  </span>
                </td>
              </tr>

              {/* Grand Total Row */}
              <tr className="bg-slate-950 font-black text-xs text-white border-t-2 border-indigo-500/40">
                <td className="p-3 uppercase text-purple-300 flex items-center gap-2">
                  <i className="fas fa-crown text-amber-400"></i> Total Hari Ini
                </td>
                <td className="p-3 text-right font-mono text-white text-sm">Rp {effectiveTotalRevenue.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-purple-300">Rp {totalActualPenyeimbang.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-emerald-300">Rp {totalActualPercepatan.toLocaleString('id-ID')}</td>
                <td className="p-3 text-right font-mono text-emerald-400 text-sm">Rp {grandTotalActualMenu.toLocaleString('id-ID')}</td>
                <td className="p-3 text-center">
                  <span className={`px-2.5 py-1 rounded-full text-[10px] font-black ${contribGrandMenuPct >= 20 ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-amber-500/20 text-amber-400 border border-amber-500/30'}`}>
                    {contribGrandMenuPct.toFixed(1)}%
                  </span>
                </td>
                <td className="p-3 text-right font-mono text-slate-300">Rp {actualGrandOutsideMenu.toLocaleString('id-ID')}</td>
                <td className="p-3 text-center text-slate-300 font-mono">{contribGrandOutsidePct.toFixed(1)}%</td>
                <td className="p-3 text-center">
                  <span className={`px-3 py-1 rounded-full text-[9px] font-black uppercase border ${grandTotalActualMenu >= grandTotalTarget ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/40' : 'bg-rose-500/20 text-rose-400 border-rose-500/40'}`}>
                    {grandTotalActualMenu >= grandTotalTarget ? 'Target Selesai' : 'Gap Perlu Dikejar'}
                  </span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* NOTES & BOTTOM BAR */}
        <div className="pt-2 border-t border-white/5 space-y-3">
          <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider">
            Catatan Pelaksanaan Strategi Hari Ini:
          </label>
          <textarea
            value={notes}
            onChange={(e) => setNotes(e.target.value)}
            rows={2}
            placeholder="Tuliskan evaluasi pelaksanaan upselling menu penyeimbang & percepatan..."
            className="w-full bg-slate-950 border border-white/10 rounded-xl p-3 text-white text-xs outline-none focus:border-purple-400"
          />

          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 pt-2">
            <span className="text-[10px] text-slate-500 font-bold">
              {storedRecord ? `Terakhir disimpan: ${new Date(storedRecord.updatedAt || '').toLocaleString('id-ID')}` : 'Belum ada data tersimpan untuk tanggal ini'}
            </span>
            <div className="flex items-center gap-3">
              <button
                type="button"
                onClick={handleCopyWhatsApp}
                className="px-4 py-3 rounded-2xl bg-emerald-700/30 hover:bg-emerald-600/40 text-emerald-300 border border-emerald-500/40 font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 cursor-pointer active:scale-95"
              >
                <i className="fab fa-whatsapp"></i> Salin WA
              </button>
              <button
                type="button"
                onClick={handleSave}
                disabled={isSaving}
                className="px-8 py-3 rounded-2xl bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-500 hover:to-indigo-500 text-white font-black text-xs uppercase tracking-wider transition-all shadow-xl shadow-purple-950/50 flex items-center gap-2 active:scale-95 cursor-pointer disabled:opacity-50"
              >
                <i className={`fas ${isSaving ? 'fa-spinner fa-spin' : 'fa-save'}`}></i>
                <span>{isSaving ? 'Menyimpan...' : 'Simpan Data Strategi'}</span>
              </button>
            </div>
          </div>
        </div>

      </div>

    </div>
  );
};

export default StrategiMenuModule;
