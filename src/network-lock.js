(() => {
  'use strict';
  const message = 'Document QC blocks runtime network access by design.';
  const reject = () => Promise.reject(new Error(message));
  const fail = () => { throw new Error(message); };

  try { Object.defineProperty(window, 'fetch', { value: reject, configurable: false, writable: false }); } catch (_) {}
  try { XMLHttpRequest.prototype.open = fail; } catch (_) {}
  try { Object.defineProperty(window, 'WebSocket', { value: class { constructor() { fail(); } }, configurable: false }); } catch (_) {}
  try { Object.defineProperty(window, 'EventSource', { value: class { constructor() { fail(); } }, configurable: false }); } catch (_) {}
  try { Object.defineProperty(navigator, 'sendBeacon', { value: () => false, configurable: false }); } catch (_) {}

  Object.defineProperty(window, '__DOCUMENT_QC_NETWORK_LOCKED__', {
    value: true,
    writable: false,
    configurable: false
  });
})();
