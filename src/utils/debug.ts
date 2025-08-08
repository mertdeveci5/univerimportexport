/**
 * Debug logging utility that can be disabled in production
 * Set DEBUG environment variable or window.__DEBUG__ to enable logging
 */

const isDebugEnabled = (): boolean => {
  // Check if we're in Node.js environment
  if (typeof process !== 'undefined' && process.env) {
    return process.env.NODE_ENV === 'development' || process.env.DEBUG === 'true';
  }
  
  // Check if we're in browser environment
  if (typeof window !== 'undefined') {
    return (window as any).__DEBUG__ === true;
  }
  
  // Default to disabled in production
  return false;
};

const noop = () => {};

export const debug = {
  log: isDebugEnabled() ? console.log.bind(console) : noop,
  warn: isDebugEnabled() ? console.warn.bind(console) : noop,
  error: console.error.bind(console), // Always keep error logging
  info: isDebugEnabled() ? console.info.bind(console) : noop,
  debug: isDebugEnabled() ? console.debug.bind(console) : noop,
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