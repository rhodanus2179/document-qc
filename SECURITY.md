# Security model

## Core privacy invariant

Selected `.docx` and `.xlsx` file contents must never be transmitted by the application.

The application therefore:

1. has no backend endpoint;
2. uses browser `File` / `ArrayBuffer` APIs only;
3. declares `connect-src 'none'` in Content Security Policy;
4. disables `fetch`, XMLHttpRequest, WebSocket, EventSource and `navigator.sendBeacon` at runtime;
5. does not load third-party JavaScript, CSS, fonts or analytics;
6. does not persist parsed document content to localStorage, IndexedDB or Cache Storage.

## What this does not guarantee

This project controls its own application code. It cannot guarantee the behavior of the browser, browser extensions, operating system, endpoint security software, or other software on the device. Users must still follow their organization's information-security policy.

## Review checklist for changes

Any pull request that introduces one of the following requires explicit security review:

- `fetch`, XMLHttpRequest, WebSocket, EventSource, `sendBeacon`
- `<form>` submission
- third-party scripts or CDN dependencies
- analytics / error-reporting SDKs
- service workers
- localStorage / IndexedDB persistence of document-derived data
- cloud AI APIs or any remote model fallback

## Reporting

Use GitHub's private vulnerability reporting feature if enabled for the repository. Do not post sensitive document samples in public issues.
