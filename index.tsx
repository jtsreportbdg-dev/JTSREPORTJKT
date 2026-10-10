
import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';

// Prevent unhandled Firestore SDK internal assertion errors (ca9, b815) from breaking the UI
const isFirestoreAssertion = (err: any) => {
  const str = String(err?.message || err?.reason?.message || err || '');
  return str.includes('INTERNAL ASSERTION FAILED') || str.includes('Unexpected state') || str.includes('(ID: ca9)') || str.includes('(ID: b815)');
};

window.addEventListener('error', (event) => {
  if (isFirestoreAssertion(event.error) || isFirestoreAssertion(event.message)) {
    console.warn('[Firestore SDK Assertion Suppressed]:', event.message || event.error);
    event.preventDefault();
    event.stopImmediatePropagation();
    return true;
  }
}, true);

window.addEventListener('unhandledrejection', (event) => {
  if (isFirestoreAssertion(event.reason)) {
    console.warn('[Firestore SDK Rejection Suppressed]:', event.reason);
    event.preventDefault();
    event.stopImmediatePropagation();
  }
}, true);

const rootElement = document.getElementById('root');
if (!rootElement) {
  throw new Error("Could not find root element to mount to");
}

const root = ReactDOM.createRoot(rootElement);
root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
