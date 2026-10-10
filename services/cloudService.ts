import { 
  doc, 
  getDoc, 
  setDoc, 
  updateDoc,
  onSnapshot,
  getDocFromServer,
  getDocs,
  getDocsFromServer,
  Timestamp,
  collection,
  addDoc,
  query,
  orderBy,
  limit,
  serverTimestamp,
  deleteField
} from "firebase/firestore";
import { auth, db } from "../firebase";
import { safeStorage } from "../lib/safeStorage";

import { Notification, Outlet } from "../types";
import { OUTLETS } from "../constants";

export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
  LISTEN = 'listen'
}

export interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  }
}

const COLLECTION_NAME = "dashboards";
const GLOBAL_COLLECTION = "system_data";
const GLOBAL_DOC_ID = "global_data";
const GLOBAL_REPORTS_ID = "global_reports";
const NOTIFICATIONS_COLLECTION = "notifications";

// Helper to remove undefined values recursively
function cleanData(obj: any): any {
  if (Array.isArray(obj)) {
    return obj.map(v => cleanData(v));
  } else if (obj !== null && typeof obj === 'object') {
    // Skip cleaning for Firestore special objects or common sentinel values
    if (obj instanceof Timestamp) return obj;
    
    // Check for FieldValue (serverTimestamp, arrayUnion, etc)
    // FieldValue instances in JS SDK v9+ often have a specific internal structure 
    // or constructor name. 'FieldValueImpl' is common in some builds.
    // Also check for _methodName which is used by sentinels.
    const constructorName = obj.constructor ? obj.constructor.name : '';
    if (
      constructorName.includes('FieldValue') || 
      obj._methodName || 
      obj.type === 'serverTimestamp' ||
      (obj._delegate && typeof obj._delegate === 'object')
    ) {
      return obj;
    }

    // Only clean plain objects to avoid breaking custom class instances
    const isPlain = obj.constructor === Object || Object.getPrototypeOf(obj) === null;
    if (!isPlain) return obj;

    return Object.fromEntries(
      Object.entries(obj)
        .filter(([_, v]) => v !== undefined)
        .map(([k, v]) => [k, cleanData(v)])
    );
  }
  return obj;
}

export function mergeDeep(local: any, cloud: any): any {
  if (local === undefined) return cloud;
  if (cloud === undefined) return local;
  if (local === null || cloud === null) return cloud !== undefined ? cloud : local;
  if (typeof local !== 'object' || typeof cloud !== 'object') return cloud;
  if (Array.isArray(local) || Array.isArray(cloud)) return cloud || local;
  
  const merged: any = { ...local };
  Object.keys(cloud).forEach(key => {
    if (local[key] !== undefined) {
      merged[key] = mergeDeep(local[key], cloud[key]);
    } else {
      merged[key] = cloud[key];
    }
  });
  return merged;
}

// Clear any persistent stale quota exceed settings on load to make the app online again
try {
  safeStorage.removeItem('FIRESTORE_QUOTA_EXCEEDED');
  safeStorage.removeItem('FIRESTORE_QUOTA_ERROR_TIME');
} catch (e) {
  // Ignore safeStorage access errors
}

let quotaExceeded = false;
let lastQuotaErrorTime = 0;
const QUOTA_RETRY_INTERVAL = 15 * 1000; // 15 seconds fast retry instead of 1 hour lockout

export function isOfflineError(error: unknown): boolean {
  const errorMessage = error instanceof Error ? error.message : String(error);
  return (
    errorMessage.includes('client is offline') || 
    errorMessage.includes('Failed to get document because the client is offline') ||
    errorMessage.includes('offline') ||
    errorMessage.includes('Could not reach Cloud Firestore')
  );
}

function checkQuota(): boolean {
  return true;
}

export async function pushNotification(user: string, outlet: Outlet, message: string): Promise<boolean> {
  if (!checkQuota()) return false;
  try {
    const notifRef = collection(db, NOTIFICATIONS_COLLECTION);
    await addDoc(notifRef, {
      user,
      outlet,
      message,
      timestamp: serverTimestamp()
    });
    return true;
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, NOTIFICATIONS_COLLECTION);
    return false;
  }
}

let activeListeners: (() => void)[] = [];

export function handleFirestoreError(error: unknown, operationType: OperationType | string, path: string | null) {
  const errorMessage = error instanceof Error ? error.message : String(error);
  const isQuotaError = errorMessage.includes('resource-exhausted') || errorMessage.includes('Quota exceeded');
  
  if (isQuotaError) {
    quotaExceeded = true;
    lastQuotaErrorTime = Date.now();
    console.warn('[CloudSync] Firestore rate limit or quota exceeded. Writes will retry in 15 seconds. Active listeners kept alive.');
    return;
  }

  if (isOfflineError(error) || errorMessage.includes('INTERNAL ASSERTION FAILED') || errorMessage.includes('Unexpected state')) {
    console.warn(`[CloudSync] Firestore operation (${operationType}) noted offline or internal assertion at path (${path}):`, errorMessage);
    // Return gracefully for ALL operations during offline state or internal assertion instead of throwing fatal crashes
    return;
  }

  const errInfo: FirestoreErrorInfo = {
    error: errorMessage,
    operationType: operationType as OperationType,
    path,
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: (auth.currentUser as any)?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    }
  };
  
  const jsonError = JSON.stringify(errInfo);
  console.error('Firestore Error: ', jsonError);
  if (operationType === 'listen' || operationType === 'get') {
    console.warn(`[CloudSync] Non-fatal ${operationType} failure at path: ${path}`);
    return;
  }
  throw new Error(jsonError);
}

// CRITICAL: Call getDocFromServer to test connection on boot
export async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
    console.info('[CloudSync] Firebase connection test succeeded. Clearing quota exceeded state if any.');
    resetQuotaExceeded();
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    const isQuotaError = errorMessage.includes('resource-exhausted') || errorMessage.includes('Quota exceeded');
    if (isQuotaError) {
      quotaExceeded = true;
      lastQuotaErrorTime = Date.now();
      try {
        safeStorage.setItem('FIRESTORE_QUOTA_EXCEEDED', 'true');
        safeStorage.setItem('FIRESTORE_QUOTA_ERROR_TIME', String(lastQuotaErrorTime));
      } catch (e) {
        // Ignore safeStorage access errors
      }
    } else if (isOfflineError(error) || errorMessage.includes('INTERNAL ASSERTION FAILED') || errorMessage.includes('Unexpected state')) {
      console.warn("[CloudSync] Client is offline or assertion:", errorMessage);
    } else {
      console.warn("[CloudSync] Connection test warning:", errorMessage);
    }
  }
}
// Removed top-level testConnection() to avoid race conditions

// Global console filter to suppress persistent Firestore quota messages
const originalError = console.error;
const originalWarn = console.warn;

console.error = (...args) => {
  const msg = args.join(' ');
  if (
    msg.includes('resource-exhausted') || 
    msg.includes('Quota exceeded') ||
    msg.includes('client is offline') ||
    msg.includes('Client is offline') ||
    msg.includes('Failed to get document because the client is offline')
  ) return;
  originalError.apply(console, args);
};

console.warn = (...args) => {
  const msg = args.join(' ');
  if (msg.includes('maximum backoff delay')) return;
  originalWarn.apply(console, args);
};

export function startNotificationListener(onUpdate: (notifications: Notification[]) => void) {
  if (!checkQuota()) return () => {};
  
  const notifRef = collection(db, NOTIFICATIONS_COLLECTION);
  const q = query(notifRef, orderBy("timestamp", "desc"), limit(10));
  
  const unsub = onSnapshot(q, (snap) => {
    const notifs: Notification[] = [];
    snap.forEach((doc) => {
      const data = doc.data();
      notifs.push({
        id: doc.id,
        user: data.user,
        outlet: data.outlet,
        message: data.message,
        timestamp: data.timestamp?.toDate ? data.timestamp.toDate().toISOString() : (typeof data.timestamp === 'string' ? data.timestamp : new Date().toISOString())
      });
    });
    onUpdate(notifs);
  }, (error) => {
    handleFirestoreError(error, OperationType.LISTEN, NOTIFICATIONS_COLLECTION);
  });

  activeListeners.push(unsub);
  return () => {
    unsub();
    activeListeners = activeListeners.filter(u => u !== unsub);
  };
}

export async function saveNotification(notification: any): Promise<boolean> {
  if (!checkQuota()) return false;
  try {
    const notifRef = collection(db, NOTIFICATIONS_COLLECTION);
    const payload = {
      ...notification,
      timestamp: typeof notification.timestamp === 'string' ? serverTimestamp() : notification.timestamp
    };
    await addDoc(notifRef, payload);
    return true;
  } catch (error) {
    console.error('Error saving notification:', error);
    return false;
  }
}

export function getFlatUpdates(newValue: any, oldValue: any, prefix = "", depth = 0): Record<string, any> {
  const updates: Record<string, any> = {};

  if (!newValue || typeof newValue !== 'object' || Array.isArray(newValue)) {
    updates[prefix] = newValue;
    return updates;
  }

  // To prevent invalid Firestore field paths when keys contain special characters or dots,
  // and to avoid excessive fragmentation, stop flattening after depth 3.
  if (depth >= 3) {
    updates[prefix] = newValue;
    return updates;
  }

  Object.keys(newValue).forEach(key => {
    // If the key itself contains dots or special characters, don't flatten inside it to prevent FieldPath errors
    const hasDot = key.includes('.');
    const fullKey = prefix ? `${prefix}.${key}` : key;
    const newVal = newValue[key];
    const oldVal = oldValue ? oldValue[key] : undefined;

    if (newVal === undefined) return;

    if (hasDot || depth >= 2) {
      if (oldVal === undefined || JSON.stringify(newVal) !== JSON.stringify(oldVal)) {
        updates[fullKey] = newVal;
      }
      return;
    }

    if (oldVal === undefined) {
      updates[fullKey] = newVal;
    } else if (JSON.stringify(newVal) !== JSON.stringify(oldVal)) {
      if (newVal && typeof newVal === 'object' && !Array.isArray(newVal)) {
        if (oldVal && typeof oldVal === 'object' && !Array.isArray(oldVal)) {
          const subUpdates = getFlatUpdates(newVal, oldVal, fullKey, depth + 1);
          Object.assign(updates, subUpdates);
        } else {
          updates[fullKey] = newVal;
        }
      } else {
        updates[fullKey] = newVal;
      }
    }
  });

  return updates;
}

export function unpackOutletDocument(docData: any): any {
  if (!docData) return {};
  const rawMonthlyData = docData.monthlyData || {};
  const monthlyData: Record<string, any> = {};
  
  // Strictly only process valid numeric month keys 0..11 to avoid key pollution and data explosion
  for (let m = 0; m < 12; m++) {
    const k = m.toString();
    if (rawMonthlyData[k] && typeof rawMonthlyData[k] === 'object') {
      const { monthlyData: _ignored, ...cleanMonth } = rawMonthlyData[k];
      monthlyData[k] = cleanMonth;
    }
  }

  const rootStaffList = Array.isArray(docData.staffList) ? docData.staffList : [];
  const rootLeaveRecords = Array.isArray(docData.leaveRecords) ? docData.leaveRecords : [];
  const rootManpowerData = (docData.manpowerData && typeof docData.manpowerData === 'object') ? docData.manpowerData : {};
  const rootDailyOvertime = (docData.dailyOvertime && typeof docData.dailyOvertime === 'object') ? docData.dailyOvertime : {};
  const rootDailyAttendance = (docData.dailyAttendance && typeof docData.dailyAttendance === 'object') ? docData.dailyAttendance : {};
  const rootDailyExternalAudit = (docData.dailyExternalAudit && typeof docData.dailyExternalAudit === 'object') ? docData.dailyExternalAudit : {};
  const rootDailyCPA = (docData.dailyCPA && typeof docData.dailyCPA === 'object') ? docData.dailyCPA : {};
  const rootDailyGSI = (docData.dailyGSI && typeof docData.dailyGSI === 'object') ? docData.dailyGSI : {};
  const rootGuestComplaints = Array.isArray(docData.guestComplaints) ? docData.guestComplaints : [];
  const rootAffiliateHistory = Array.isArray(docData.affiliateHistory) ? docData.affiliateHistory : [];

  // Consolidate bestStaffList across root and all 12 months
  let bestStaffList = rootStaffList;
  if (bestStaffList.length === 0) {
    for (let m = 11; m >= 0; m--) {
      const mStaff = monthlyData[m.toString()]?.staffList;
      if (Array.isArray(mStaff) && mStaff.length > 0) {
        bestStaffList = mStaff;
        break;
      }
    }
    if (bestStaffList.length === 0) {
      for (let m = 0; m < 12; m++) {
        const mStaff = monthlyData[m.toString()]?.staffList;
        if (Array.isArray(mStaff) && mStaff.length > 0) {
          bestStaffList = mStaff;
          break;
        }
      }
    }
  }

  // Consolidate bestLeaveRecords across root and all 12 months
  let bestLeaveRecords = rootLeaveRecords;
  if (bestLeaveRecords.length === 0) {
    for (let m = 11; m >= 0; m--) {
      const mLeave = monthlyData[m.toString()]?.leaveRecords;
      if (Array.isArray(mLeave) && mLeave.length > 0) {
        bestLeaveRecords = mLeave;
        break;
      }
    }
    if (bestLeaveRecords.length === 0) {
      for (let m = 0; m < 12; m++) {
        const mLeave = monthlyData[m.toString()]?.leaveRecords;
        if (Array.isArray(mLeave) && mLeave.length > 0) {
          bestLeaveRecords = mLeave;
          break;
        }
      }
    }
  }

  // Consolidate manpowerData across root and all 12 months
  const consolidatedManpowerData: Record<string, any> = { ...rootManpowerData };
  const consolidatedDailyOvertime: Record<string, any[]> = { ...rootDailyOvertime };
  const consolidatedDailyAttendance: Record<string, any[]> = { ...rootDailyAttendance };
  const consolidatedDailyExternalAudit: Record<string, any[]> = { ...rootDailyExternalAudit };
  const consolidatedDailyCPA: Record<string, any[]> = { ...rootDailyCPA };
  const consolidatedDailyGSI: Record<string, any[]> = { ...rootDailyGSI };
  const consolidatedGuestComplaints: any[] = [...rootGuestComplaints];
  const consolidatedAffiliateHistory: any[] = [...rootAffiliateHistory];

  for (let m = 0; m < 12; m++) {
    const mData = monthlyData[m.toString()];
    if (!mData) continue;

    if (mData.manpowerData && typeof mData.manpowerData === 'object') {
      Object.assign(consolidatedManpowerData, mData.manpowerData);
    }
    if (mData.dailyOvertime && typeof mData.dailyOvertime === 'object') {
      Object.assign(consolidatedDailyOvertime, mData.dailyOvertime);
    }
    if (mData.dailyAttendance && typeof mData.dailyAttendance === 'object') {
      Object.assign(consolidatedDailyAttendance, mData.dailyAttendance);
    }
    if (mData.dailyExternalAudit && typeof mData.dailyExternalAudit === 'object') {
      Object.assign(consolidatedDailyExternalAudit, mData.dailyExternalAudit);
    }
    if (mData.dailyCPA && typeof mData.dailyCPA === 'object') {
      Object.assign(consolidatedDailyCPA, mData.dailyCPA);
    }
    if (mData.dailyGSI && typeof mData.dailyGSI === 'object') {
      Object.assign(consolidatedDailyGSI, mData.dailyGSI);
    }
    if (Array.isArray(mData.guestComplaints)) {
      mData.guestComplaints.forEach((c: any) => {
        if (c && !consolidatedGuestComplaints.some(gc => gc.id === c.id)) {
          consolidatedGuestComplaints.push(c);
        }
      });
    }
    if (Array.isArray(mData.affiliateHistory)) {
      mData.affiliateHistory.forEach((a: any) => {
        if (a && !consolidatedAffiliateHistory.some(ga => ga.id === a.id)) {
          consolidatedAffiliateHistory.push(a);
        }
      });
    }
  }

  const filledMonthlyData: any = {};
  for (let m = 0; m < 12; m++) {
    const mKey = m.toString();
    const mData = monthlyData[mKey] || {};
    
    const staffListToUse = (Array.isArray(mData.staffList) && mData.staffList.length > 0)
      ? mData.staffList
      : bestStaffList;

    const leaveRecordsToUse = (Array.isArray(mData.leaveRecords) && mData.leaveRecords.length > 0)
      ? mData.leaveRecords
      : bestLeaveRecords;

    const manpowerDataToUse = {
      ...consolidatedManpowerData,
      ...(mData.manpowerData && typeof mData.manpowerData === 'object' ? mData.manpowerData : {})
    };

    filledMonthlyData[mKey] = {
      ...mData,
      staffList: staffListToUse,
      leaveRecords: leaveRecordsToUse,
      manpowerData: manpowerDataToUse,
      dailyOvertime: { ...consolidatedDailyOvertime, ...(mData.dailyOvertime || {}) },
      dailyAttendance: { ...consolidatedDailyAttendance, ...(mData.dailyAttendance || {}) },
      dailyExternalAudit: { ...consolidatedDailyExternalAudit, ...(mData.dailyExternalAudit || {}) },
      dailyCPA: { ...consolidatedDailyCPA, ...(mData.dailyCPA || {}) },
      dailyGSI: { ...consolidatedDailyGSI, ...(mData.dailyGSI || {}) },
      guestComplaints: consolidatedGuestComplaints,
      affiliateHistory: consolidatedAffiliateHistory
    };
  }

  // Attach monthlyData map for components that inspect outletMonths.monthlyData?.[m]
  filledMonthlyData.monthlyData = monthlyData;

  // Consolidate logistics across months and root
  const consolidatedLogisticItems: any[] = Array.isArray(docData.logisticItems) ? [...docData.logisticItems] : [];
  const consolidatedDailyLogisticRecords: Record<string, any> = { ...(docData.dailyLogisticRecords || {}) };
  const consolidatedPeminjamanRecords: any[] = Array.isArray(docData.peminjamanRecords) ? [...docData.peminjamanRecords] : [];
  const consolidatedLogisticBeginningStocks: Record<string, any> = { ...(docData.logisticBeginningStocks || {}) };
  const consolidatedLogisticPeriodPrices: Record<string, any> = { ...(docData.logisticPeriodPrices || {}) };

  for (let m = 0; m < 12; m++) {
    const mData = monthlyData[m.toString()];
    if (!mData) continue;
    if (Array.isArray(mData.logisticItems) && consolidatedLogisticItems.length === 0) {
      consolidatedLogisticItems.push(...mData.logisticItems);
    }
    if (mData.dailyLogisticRecords && typeof mData.dailyLogisticRecords === 'object') {
      Object.assign(consolidatedDailyLogisticRecords, mData.dailyLogisticRecords);
    }
    if (Array.isArray(mData.peminjamanRecords)) {
      mData.peminjamanRecords.forEach((p: any) => {
        if (p && p.id && !consolidatedPeminjamanRecords.some(cp => cp.id === p.id)) {
          consolidatedPeminjamanRecords.push(p);
        }
      });
    }
    if (mData.logisticBeginningStocks && typeof mData.logisticBeginningStocks === 'object') {
      Object.assign(consolidatedLogisticBeginningStocks, mData.logisticBeginningStocks);
    }
    if (mData.logisticPeriodPrices && typeof mData.logisticPeriodPrices === 'object') {
      Object.assign(consolidatedLogisticPeriodPrices, mData.logisticPeriodPrices);
    }
  }

  // Find most recent non-empty sales & expenses to attach as default root objects
  let defaultSales = docData.sales || null;
  let defaultExpenses = docData.expenses || null;
  for (let m = 11; m >= 0; m--) {
    const mData = monthlyData[m.toString()];
    if (!defaultSales && mData?.sales) defaultSales = mData.sales;
    if (!defaultExpenses && mData?.expenses) defaultExpenses = mData.expenses;
  }
  if (defaultSales) {
    filledMonthlyData.sales = defaultSales;
    if (defaultSales.dailyRecording) {
      filledMonthlyData.dailyRecording = defaultSales.dailyRecording;
    }
  }
  if (defaultExpenses) filledMonthlyData.expenses = defaultExpenses;
  if (docData.monthlyCogsData) filledMonthlyData.monthlyCogsData = docData.monthlyCogsData;
  if (docData.dailyTargets) filledMonthlyData.dailyTargets = docData.dailyTargets;
  if (docData.dailyQA) filledMonthlyData.dailyQA = docData.dailyQA;
  if (docData.dailyMainKitchenLedger) filledMonthlyData.dailyMainKitchenLedger = docData.dailyMainKitchenLedger;
  if (docData.dailyMainStoreLedger) filledMonthlyData.dailyMainStoreLedger = docData.dailyMainStoreLedger;
  if (docData.dailyChemicalLedger) filledMonthlyData.dailyChemicalLedger = docData.dailyChemicalLedger;
  if (docData.dailyCashLedger) filledMonthlyData.dailyCashLedger = docData.dailyCashLedger;
  if (docData.dailyCommodityLedger) filledMonthlyData.dailyCommodityLedger = docData.dailyCommodityLedger;
  if (docData.dailyGuestSuppliesLedger) filledMonthlyData.dailyGuestSuppliesLedger = docData.dailyGuestSuppliesLedger;
  if (docData.dailyCityLedger) filledMonthlyData.dailyCityLedger = docData.dailyCityLedger;
  if (docData.dailyTgifLedger) filledMonthlyData.dailyTgifLedger = docData.dailyTgifLedger;
  if (docData.dailySukandaLedger) filledMonthlyData.dailySukandaLedger = docData.dailySukandaLedger;
  if (docData.dailyEverestLedger) filledMonthlyData.dailyEverestLedger = docData.dailyEverestLedger;
  if (docData.dailyPanfilaLedger) filledMonthlyData.dailyPanfilaLedger = docData.dailyPanfilaLedger;
  if (docData.dailyIkigaiLedger) filledMonthlyData.dailyIkigaiLedger = docData.dailyIkigaiLedger;
  if (docData.dailyKebunkuLedger) filledMonthlyData.dailyKebunkuLedger = docData.dailyKebunkuLedger;
  if (docData.dailySaungKubilLedger) filledMonthlyData.dailySaungKubilLedger = docData.dailySaungKubilLedger;
  if (docData.dailyGasLedger) filledMonthlyData.dailyGasLedger = docData.dailyGasLedger;
  if (docData.dailyThirdPartyLedger) filledMonthlyData.dailyThirdPartyLedger = docData.dailyThirdPartyLedger;
  if (docData.energyCost) filledMonthlyData.energyCost = docData.energyCost;
  if (docData.weeklyOvertime) filledMonthlyData.weeklyOvertime = docData.weeklyOvertime;
  if (docData.dailyBreakTime) filledMonthlyData.dailyBreakTime = docData.dailyBreakTime;
  if (docData.googleReview) filledMonthlyData.googleReview = docData.googleReview;
  if (docData.dailyGoogleReview) filledMonthlyData.dailyGoogleReview = docData.dailyGoogleReview;
  if (docData.promoSettings) filledMonthlyData.promoSettings = docData.promoSettings;
  if (docData.bankingPromos) filledMonthlyData.bankingPromos = docData.bankingPromos;
  if (docData.dailyStrategy) filledMonthlyData.dailyStrategy = docData.dailyStrategy;

  // Consolidate dailyStrategy across months
  const consolidatedDailyStrategy: Record<string, any> = { ...(docData.dailyStrategy || {}) };
  for (let m = 0; m < 12; m++) {
    const mData = monthlyData[m.toString()];
    if (mData?.dailyStrategy && typeof mData.dailyStrategy === 'object') {
      Object.assign(consolidatedDailyStrategy, mData.dailyStrategy);
    }
  }
  filledMonthlyData.dailyStrategy = consolidatedDailyStrategy;

  filledMonthlyData.logisticItems = consolidatedLogisticItems;
  filledMonthlyData.dailyLogisticRecords = consolidatedDailyLogisticRecords;
  filledMonthlyData.peminjamanRecords = consolidatedPeminjamanRecords;
  filledMonthlyData.logisticBeginningStocks = consolidatedLogisticBeginningStocks;
  filledMonthlyData.logisticPeriodPrices = consolidatedLogisticPeriodPrices;

  return filledMonthlyData;
}

export async function saveGlobalData(email: string, data: Partial<any>, oldData?: Partial<any>, outletId?: string): Promise<boolean> {
  if (outletId) {
    return saveOutletData(email, outletId, {}, undefined, data);
  }
  return true;
}

let lastBackupTimeByOutlet: Record<string, number> = {};

// Throttling and batching Firestore writes per outlet to protect Firestore daily quota limits
const lastFirestoreWriteByOutlet: Record<string, number> = {};
const pendingFirestoreTimerByOutlet: Record<string, any> = {};
const pendingFirestorePayloadByOutlet: Record<string, { email: string; data: Record<string, any>; extraGlobal?: Record<string, any> }> = {};
const FIRESTORE_MIN_WRITE_INTERVAL_MS = 5000; // at most 1 write every 5 seconds per outlet to protect daily write quota

async function executeFirestoreWrite(
  email: string,
  outletId: string,
  dataToSave: Record<string, any>,
  extraGlobalToSave?: Record<string, any>
): Promise<boolean> {
  if (!checkQuota()) return false;
  try {
    const outletRef = doc(db, COLLECTION_NAME, outletId);
    const completeMonthsToMerge: Record<string, any> = {};
    const payloadToMerge: Record<string, any> = {
      updatedAt: serverTimestamp(),
      lastUpdatedBy: (email || 'outlet-sync').toLowerCase(),
      isDeleted: false
    };

    let hasData = false;
    Object.entries(dataToSave).forEach(([monthIdx, monthData]) => {
      const mNum = parseInt(monthIdx, 10);
      if (isNaN(mNum) || mNum < 0 || mNum > 11 || !monthData) return;
      const { monthlyData: _ignored, ...cleanMonth } = cleanData(monthData);
      completeMonthsToMerge[monthIdx] = cleanMonth;
      hasData = true;
    });

    if (Object.keys(completeMonthsToMerge).length > 0) {
      payloadToMerge.monthlyData = completeMonthsToMerge;
    }

    let topStaffList: any[] | null = null;
    let topLeaveRecords: any[] | null = null;
    const topManpowerData: Record<string, any> = {};

    Object.values(dataToSave).forEach((monthData: any) => {
      if (monthData) {
        if (Array.isArray(monthData.staffList) && monthData.staffList.length > 0) {
          topStaffList = monthData.staffList;
        }
        if (Array.isArray(monthData.leaveRecords) && monthData.leaveRecords.length > 0) {
          topLeaveRecords = monthData.leaveRecords;
        }
        if (monthData.manpowerData && typeof monthData.manpowerData === 'object') {
          Object.assign(topManpowerData, monthData.manpowerData);
        }
      }
    });

    if (topStaffList !== null) { payloadToMerge.staffList = topStaffList; hasData = true; }
    if (topLeaveRecords !== null) { payloadToMerge.leaveRecords = topLeaveRecords; hasData = true; }
    if (Object.keys(topManpowerData).length > 0) { payloadToMerge.manpowerData = topManpowerData; hasData = true; }

    Object.values(dataToSave).forEach((mVal: any) => {
      if (mVal) {
        const topKeys = [
          'sales', 'dailyTargets', 'expenses', 'monthlyCogsData',
          'dailyMainKitchenLedger', 'dailyMainStoreLedger', 'dailyChemicalLedger',
          'dailyCashLedger', 'dailyCommodityLedger', 'dailyGuestSuppliesLedger',
          'dailyCityLedger', 'dailyQA', 'dailyGoogleReview', 'dailyBreakTime',
          'dailyTgifLedger', 'dailySukandaLedger', 'dailyEverestLedger',
          'dailyPanfilaLedger', 'dailyIkigaiLedger', 'dailyKebunkuLedger',
          'dailySaungKubilLedger', 'dailyGasLedger', 'dailyThirdPartyLedger',
          'energyCost', 'weeklyOvertime', 'staffList', 'leaveRecords', 'manpowerData',
          'promoSettings', 'bankingPromos', 'dailyStrategy'
        ];
        topKeys.forEach(k => {
          if (mVal[k] !== undefined) {
            payloadToMerge[k] = cleanData(mVal[k]);
            hasData = true;
          }
        });
      }
    });

    if (extraGlobalToSave && typeof extraGlobalToSave === 'object') {
      const allowedGlobals = [
        'staffList', 'leaveRecords', 'dailyOvertime', 'weeklyOvertime', 'dailyAttendance',
        'dailyExternalAudit', 'dailyCPA', 'dailyGSI', 'guestComplaints',
        'affiliateHistory', 'manpowerData'
      ];
      allowedGlobals.forEach(gKey => {
        if (extraGlobalToSave[gKey] !== undefined) {
          payloadToMerge[gKey] = extraGlobalToSave[gKey];
          hasData = true;
        }
      });
    }

    if (!hasData) return true;

    // Direct atomic setDoc with merge: true - NO redundant getDoc read, saves 50% quota!
    await setDoc(outletRef, cleanData(payloadToMerge), { merge: true });
    lastFirestoreWriteByOutlet[outletId] = Date.now();
    console.log(`[CloudSync] Outlet data (${outletId}) successfully synced to Firestore.`);
    return true;
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, `${COLLECTION_NAME}/${outletId}`);
    return false;
  }
}

export async function saveOutletData(
  email: string, 
  outletId: string, 
  data: Record<string, any>, 
  oldData?: Record<string, any>,
  extraGlobal?: Record<string, any>
): Promise<boolean> {
  // 1. Immediately save & synchronize to High-Availability Cloud Server API (instant cross-device sync)
  let serverSyncSuccess = false;
  try {
    const res = await fetch(`/api/sync/outlet/${encodeURIComponent(outletId)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data, email, extraGlobal })
    });
    if (res.ok) {
      serverSyncSuccess = true;
      console.log(`[ServerSync] Outlet data (${outletId}) saved to cloud server for instant cross-device sync.`);
    }
  } catch (err) {
    console.warn('[ServerSync] Server endpoint notice:', err);
  }

  // 2. Throttle and batch Firestore writes to protect daily token/quota
  const now = Date.now();
  const lastWrite = lastFirestoreWriteByOutlet[outletId] || 0;
  const timeSinceLast = now - lastWrite;

  // Accumulate pending payload
  if (!pendingFirestorePayloadByOutlet[outletId]) {
    pendingFirestorePayloadByOutlet[outletId] = { email, data: {}, extraGlobal: {} };
  }
  pendingFirestorePayloadByOutlet[outletId].email = email;
  pendingFirestorePayloadByOutlet[outletId].data = {
    ...pendingFirestorePayloadByOutlet[outletId].data,
    ...data
  };
  if (extraGlobal) {
    pendingFirestorePayloadByOutlet[outletId].extraGlobal = {
      ...pendingFirestorePayloadByOutlet[outletId].extraGlobal,
      ...extraGlobal
    };
  }

  let firestoreSuccess = false;
  if (timeSinceLast >= FIRESTORE_MIN_WRITE_INTERVAL_MS) {
    if (pendingFirestoreTimerByOutlet[outletId]) {
      clearTimeout(pendingFirestoreTimerByOutlet[outletId]);
      delete pendingFirestoreTimerByOutlet[outletId];
    }
    const payloadToExecute = pendingFirestorePayloadByOutlet[outletId];
    delete pendingFirestorePayloadByOutlet[outletId];
    firestoreSuccess = await executeFirestoreWrite(
      payloadToExecute.email,
      outletId,
      payloadToExecute.data,
      payloadToExecute.extraGlobal
    );
  } else {
    // Within throttle window: schedule trailing write to execute once window expires
    if (!pendingFirestoreTimerByOutlet[outletId]) {
      const delay = FIRESTORE_MIN_WRITE_INTERVAL_MS - timeSinceLast;
      pendingFirestoreTimerByOutlet[outletId] = setTimeout(async () => {
        delete pendingFirestoreTimerByOutlet[outletId];
        const pending = pendingFirestorePayloadByOutlet[outletId];
        if (pending) {
          delete pendingFirestorePayloadByOutlet[outletId];
          await executeFirestoreWrite(pending.email, outletId, pending.data, pending.extraGlobal);
        }
      }, delay);
    }
    // Server sync already succeeded, so return true
    firestoreSuccess = true;
  }

  // Return true if either server sync or firestore succeeded so user edits are confirmed
  return serverSyncSuccess || firestoreSuccess;
}

export async function saveUserData(email: string, data: any, outletId?: string): Promise<boolean> {
  try {
    const globalSuccess = await saveGlobalData(email, data);
    let outletSuccess = true;
    if (outletId) {
      const { staffList, leaveRecords, dailyOvertime, dailyMeetings, ...outletData } = data;
      outletSuccess = await saveOutletData(email, outletId, outletData);
    }
    return globalSuccess && outletSuccess;
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, outletId ? `${COLLECTION_NAME}/${outletId}` : COLLECTION_NAME);
    return false;
  }
}

export async function fetchUserData(email: string, outletId?: string): Promise<any | null> {
  const cacheKey = `BARN_OFFLINE_CACHE_${email.toLowerCase()}_${outletId || 'ALL'}`;
  let combinedData: any = {};

  // 1. Fetch from high-speed Server Sync API (instant across devices)
  if (outletId) {
    try {
      const sRes = await fetch(`/api/sync/outlet/${encodeURIComponent(outletId)}`);
      if (sRes.ok) {
        const sJson = await sRes.json();
        if (sJson.outlet) {
          combinedData = unpackOutletDocument(sJson.outlet);
        }
      }
    } catch (e) {
      console.warn('[ServerSync] Server fetchUserData notice:', e);
    }
  }

  // 2. Fetch from Firestore if quota allows and merge
  if (outletId && checkQuota()) {
    try {
      const outletRef = doc(db, COLLECTION_NAME, outletId);
      let outletSnap;
      try {
        outletSnap = await getDocFromServer(outletRef);
      } catch {
        outletSnap = await getDoc(outletRef);
      }
      if (outletSnap && outletSnap.exists()) {
        const cloudData = unpackOutletDocument(outletSnap.data());
        combinedData = mergeDeep(combinedData, cloudData);
      }
    } catch (e) {
      if (!isOfflineError(e)) {
        handleFirestoreError(e, OperationType.GET, `${COLLECTION_NAME}/${outletId}`);
      }
    }
  }

  if (Object.keys(combinedData).length === 0) {
    const cached = safeStorage.getItem(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {}
    }
    return {};
  }

  // Save to local storage for future offline access
  try {
    safeStorage.setItem(cacheKey, JSON.stringify(combinedData));
  } catch (cacheErr) {}

  return combinedData;
}

export async function fetchAllOutletsData(): Promise<Record<string, any> | null> {
  const cacheKey = "BARN_OFFLINE_CACHE_ALL_OUTLETS";
  const results: Record<string, any> = {};

  // 1. Fetch from server-side sync API
  try {
    const sRes = await fetch('/api/sync/all-outlets');
    if (sRes.ok) {
      const sJson = await sRes.json();
      if (sJson.outlets && typeof sJson.outlets === 'object') {
        Object.entries(sJson.outlets).forEach(([id, data]: [string, any]) => {
          if (!data) return;
          const filled = unpackOutletDocument(data);
          results[id] = {
            ...filled,
            isDeleted: false,
            _googleReview: data.googleReview || null,
            _lastUpdated: data.updatedAt || null
          };
        });
      }
    }
  } catch (err) {
    console.warn('[ServerSync] Fetch all outlets notice:', err);
  }

  // 2. Fetch from Firestore if quota allows and merge with latest
  if (checkQuota()) {
    try {
      const colRef = collection(db, COLLECTION_NAME);
      let snap;
      try {
        snap = await getDocsFromServer(colRef);
      } catch {
        snap = await getDocs(colRef);
      }
      snap.forEach(docSnap => {
        const data = docSnap.data();
        const filled = unpackOutletDocument(data);
        const existing = results[docSnap.id];
        const cloudUpdated = data.updatedAt?.toDate?.()?.toISOString() || null;
        if (!existing || (cloudUpdated && (!existing._lastUpdated || cloudUpdated >= existing._lastUpdated))) {
          results[docSnap.id] = {
            ...filled,
            isDeleted: false,
            _googleReview: data.googleReview || null,
            _lastUpdated: cloudUpdated || existing?._lastUpdated || null
          };
        }
      });
    } catch (err) {
      handleFirestoreError(err, OperationType.GET, COLLECTION_NAME);
    }
  }

  if (Object.keys(results).length > 0) {
    try {
      safeStorage.setItem(cacheKey, JSON.stringify(results));
    } catch {}
    return results;
  }

  const cached = safeStorage.getItem(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch {}
  }
  return results;
}

export function startAllOutletsListener(onUpdate: (data: Record<string, any>) => void) {
  let isSubscribed = true;
  let lastKnownHash = '';

  // Function to poll and sync from server API
  const pollServerSync = async () => {
    if (!isSubscribed) return;
    try {
      const res = await fetch('/api/sync/all-outlets');
      if (res.ok) {
        const json = await res.json();
        if (json.outlets && typeof json.outlets === 'object') {
          const processed: Record<string, any> = {};
          Object.entries(json.outlets).forEach(([id, data]: [string, any]) => {
            if (!data) return;
            const filled = unpackOutletDocument(data);
            processed[id] = {
              ...filled,
              isDeleted: false,
              _googleReview: data.googleReview || null,
              _lastUpdated: data.updatedAt || null
            };
          });
          const newHash = JSON.stringify(processed);
          if (newHash !== lastKnownHash) {
            lastKnownHash = newHash;
            onUpdate(processed);
          }
        }
      }
    } catch (e) {
      // Ignore background poll errors
    }
  };

  pollServerSync();
  const pollInterval = setInterval(pollServerSync, 3500);

  let unsubFirestore: (() => void) | null = null;
  if (checkQuota()) {
    try {
      const colRef = collection(db, COLLECTION_NAME);
      unsubFirestore = onSnapshot(colRef, (snap) => {
        const results: any = {};
        snap.forEach((docSnap) => {
          const data = docSnap.data();
          const filled = unpackOutletDocument(data);
          results[docSnap.id] = {
            ...filled,
            isDeleted: false,
            _googleReview: data.googleReview || null,
            _lastUpdated: data.updatedAt?.toDate?.()?.toISOString() || null
          };
        });
        const newHash = JSON.stringify(results);
        if (newHash !== lastKnownHash) {
          lastKnownHash = newHash;
          onUpdate(results);
        }
      }, (error) => {
        handleFirestoreError(error, OperationType.LISTEN, COLLECTION_NAME);
      });
    } catch {}
  }

  const cleanup = () => {
    isSubscribed = false;
    clearInterval(pollInterval);
    if (unsubFirestore) unsubFirestore();
  };

  activeListeners.push(cleanup);
  return () => {
    cleanup();
    activeListeners = activeListeners.filter(u => u !== cleanup);
  };
}

export function startRealtimeListener(
  email: string, 
  outletId: string,
  onUpdate: (data: any) => void,
  onStatusChange?: (status: 'connected' | 'disconnected' | 'reconnecting') => void
) {
  let isSubscribed = true;
  let lastKnownHash = '';

  const syncFromServer = async () => {
    if (!isSubscribed) return;
    try {
      const res = await fetch(`/api/sync/outlet/${encodeURIComponent(outletId)}`);
      if (res.ok) {
        const json = await res.json();
        if (json.outlet) {
          const filled = unpackOutletDocument(json.outlet);
          const newHash = JSON.stringify(filled);
          if (newHash !== lastKnownHash) {
            lastKnownHash = newHash;
            onUpdate(filled);
            onStatusChange?.('connected');
          }
        }
      }
    } catch (e) {}
  };

  syncFromServer();
  const pollInterval = setInterval(syncFromServer, 3000);

  let unsubFirestore: (() => void) | null = null;
  if (checkQuota()) {
    try {
      const outletRef = doc(db, COLLECTION_NAME, outletId);
      unsubFirestore = onSnapshot(outletRef, (snap) => {
        if (snap.exists()) {
          const docData = snap.data();
          if (docData) {
            const filledMonthlyData = unpackOutletDocument(docData);
            const newHash = JSON.stringify(filledMonthlyData);
            if (newHash !== lastKnownHash) {
              lastKnownHash = newHash;
              onUpdate(filledMonthlyData);
            }
          }
        }
        onStatusChange?.('connected');
      }, (error) => {
        handleFirestoreError(error, OperationType.LISTEN, `${COLLECTION_NAME}/${outletId}`);
      });
    } catch {}
  }

  onStatusChange?.('connected');

  const cleanup = () => {
    isSubscribed = false;
    clearInterval(pollInterval);
    if (unsubFirestore) unsubFirestore();
  };

  activeListeners.push(cleanup);
  return () => {
    cleanup();
    activeListeners = activeListeners.filter(u => u !== cleanup);
  };
}

// Note: testConnection and quota check logic is now managed by explicit calls rather than top-level side effects
// to prevent "The operation is insecure" errors on module load.

export function resetQuotaExceeded() {
  quotaExceeded = false;
  lastQuotaErrorTime = 0;
  try {
    safeStorage.removeItem('FIRESTORE_QUOTA_EXCEEDED');
    safeStorage.removeItem('FIRESTORE_QUOTA_ERROR_TIME');
  } catch (e) {
    // Ignore safeStorage access errors
  }
  console.info('[CloudSync] Quota state reset. Sync resumed.');
}

export function isQuotaExceeded(): boolean {
  return !checkQuota();
}
