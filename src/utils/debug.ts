/**
 * Debug logging utility that can be disabled in production
 * Set DEBUG environment variable or window.__DEBUG__ to enable logging
 */

const isDebugEnabled = (): boolean => {
  // TEMPORARILY ENABLE LOGGING IN ALL ENVIRONMENTS FOR DEBUGGING
  return true;

  // Original logic (disabled for now):
  // Check if we're in Node.js environment
  // if (typeof process !== 'undefined' && process.env) {
  //   return process.env.NODE_ENV === 'development' || process.env.DEBUG === 'true';
  // }

  // Check if we're in browser environment
  // if (typeof window !== 'undefined') {
  //   return (window as any).__DEBUG__ === true;
  // }

  // Default to disabled in production
  // return false;
};

const noop = () => {};

// Create safe wrappers that always work
const safeLog = (...args: any[]) => {
  if (isDebugEnabled() && console.log) {
    console.log(...args);
  }
};

const safeWarn = (...args: any[]) => {
  if (isDebugEnabled() && console.warn) {
    console.warn(...args);
  }
};

const safeError = (...args: any[]) => {
  if (console.error) {
    console.error(...args);
  }
};

const safeInfo = (...args: any[]) => {
  if (isDebugEnabled() && console.info) {
    console.info(...args);
  }
};

const safeDebug = (...args: any[]) => {
  if (isDebugEnabled() && console.debug) {
    console.debug(...args);
  }
};

export const debug = {
  log: safeLog,
  warn: safeWarn,
  error: safeError,
  info: safeInfo,
  debug: safeDebug,
};

// For development: Enable debug logging
export const enableDebug = () => {
  if (typeof window !== 'undefined') {
    (window as any).__DEBUG__ = true;
  }
};

// For production: Disable debug logging
export const disableDebug = () => {
  if (typeof window !== 'undefined') {
    (window as any).__DEBUG__ = false;
  }
};