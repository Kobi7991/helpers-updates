// ============================================
// Lemon Squeezy License Management
// Tiers: Single User (1), Multi User (5), Studio (25). Variant IDs: 1314978, 1314998, 1315002.
// ============================================

const LEMON_SQUEEZY_ACTIVATE_URL = 'https://api.lemonsqueezy.com/v1/licenses/activate';
const LEMON_SQUEEZY_DEACTIVATE_URL = 'https://api.lemonsqueezy.com/v1/licenses/deactivate';
const LEMON_SQUEEZY_VALIDATE_URL = 'https://api.lemonsqueezy.com/v1/licenses/validate';

const LICENSE_STORAGE_KEY = 'helpers_license_verified';
const LICENSE_KEY_STORAGE_KEY = 'helpers_license_key_hash';
const LICENSE_KEY_LAST4_KEY = 'helpers_license_key_last4';
const LICENSE_SEAT_USES_KEY = 'helpers_license_seat_uses';
const LICENSE_SEAT_QUANTITY_KEY = 'helpers_license_seat_quantity';
const LICENSE_KEY_FOR_RELEASE_KEY = 'helpers_license_key_for_release';
const LICENSE_INSTANCE_ID_KEY = 'helpers_license_instance_id';
const LICENSE_TIER_NAME_KEY = 'helpers_license_tier_name';

/** Variant ID to tier name (fallback when meta.variant_name missing). */
var LEMON_VARIANT_NAMES = { 1314978: 'Single User', 1314998: 'Multi User', 1315002: 'Studio' };

/**
 * Get or create a stable instance name for this machine (so we don't burn multiple activations).
 */
function getOrCreateInstanceName() {
    var key = 'helpers_license_instance_name';
    try {
        var existing = localStorage.getItem(key);
        if (existing && existing.length > 0) return existing;
        var id = 'helpers-' + Math.random().toString(36).slice(2) + '-' + Date.now().toString(36);
        localStorage.setItem(key, id);
        return id;
    } catch (e) {
        return 'helpers-' + Date.now();
    }
}

/**
 * Verify license key with Lemon Squeezy activate endpoint.
 */
function verifyLicense(key) {
    var btn = document.getElementById('activate-btn');
    var keyInput = document.getElementById('license-key-input');

    if (!key || !key.trim()) {
        showError('Please enter a license key');
        return;
    }

    var licenseKey = key.trim();
    if (licenseKey.length < 8) {
        showError('License key is too short. Please check and try again.');
        resetButton();
        return;
    }

    hideError();
    btn.textContent = 'Verifying...';
    btn.disabled = true;
    if (keyInput) keyInput.disabled = true;

    var body = 'license_key=' + encodeURIComponent(licenseKey) +
               '&instance_name=' + encodeURIComponent(getOrCreateInstanceName());
    var xhr = new XMLHttpRequest();
    xhr.open('POST', LEMON_SQUEEZY_ACTIVATE_URL, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    xhr.setRequestHeader('Accept', 'application/json');
    xhr.timeout = 10000;
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status === 0) {
            showError('Could not reach the license server. Check your internet connection and try again.');
            resetButton();
            return;
        }
        if (xhr.status >= 400) {
            var errMsg = 'License server returned an error. Please try again or contact support.';
            try {
                var errData = JSON.parse(xhr.responseText);
                if (errData && errData.error) errMsg = errData.error;
            } catch (e) {}
            showError(errMsg);
            resetButton();
            return;
        }
        try {
            var data = JSON.parse(xhr.responseText);
            if (!data || data.activated !== true || data.error) {
                showError((data && data.error) ? data.error : 'Invalid license key. Please check and try again.');
                resetButton();
                return;
            }
            var lk = data.license_key;
            var meta = data.meta || {};
            var instance = data.instance;
            var usage = (lk && lk.activation_usage != null) ? Number(lk.activation_usage) : 0;
            var limit = (lk && lk.activation_limit != null) ? Number(lk.activation_limit) : 1;
            var tierName = (meta.variant_name && meta.variant_name.trim()) ? meta.variant_name.trim()
                : (meta.variant_id != null && LEMON_VARIANT_NAMES[meta.variant_id]) ? LEMON_VARIANT_NAMES[meta.variant_id] : 'License';
            if (instance && instance.id) {
                try { localStorage.setItem(LICENSE_INSTANCE_ID_KEY, instance.id); } catch (e) {}
            }
            saveLicenseStatus(true, licenseKey);
            saveLicenseSeats(usage, limit);
            try { localStorage.setItem(LICENSE_TIER_NAME_KEY, tierName); } catch (e) {}
            hideOverlay();
            showSuccess();
        } catch (e) {
            showError('Error parsing server response. Please try again.');
        }
        resetButton();
    };
    xhr.ontimeout = function () {
        showError('Request timed out. Please check your connection and try again.');
        resetButton();
    };
    xhr.onerror = function () {
        showError('Network error. Please check your connection and try again.');
        resetButton();
    };
    xhr.send(body);
}

/**
 * Save license verification status securely
 */
function saveLicenseStatus(verified, licenseKey) {
    try {
        if (verified) {
            localStorage.setItem(LICENSE_STORAGE_KEY, 'true');
            var hash = simpleHash(licenseKey || '');
            localStorage.setItem(LICENSE_KEY_STORAGE_KEY, hash);
            if (licenseKey && licenseKey.length >= 4) {
                localStorage.setItem(LICENSE_KEY_LAST4_KEY, licenseKey.slice(-4));
            }
            if (licenseKey) localStorage.setItem(LICENSE_KEY_FOR_RELEASE_KEY, licenseKey);
        } else {
            localStorage.removeItem(LICENSE_STORAGE_KEY);
            localStorage.removeItem(LICENSE_KEY_STORAGE_KEY);
            localStorage.removeItem(LICENSE_KEY_LAST4_KEY);
            localStorage.removeItem(LICENSE_KEY_FOR_RELEASE_KEY);
            localStorage.removeItem(LICENSE_SEAT_USES_KEY);
            localStorage.removeItem(LICENSE_SEAT_QUANTITY_KEY);
            localStorage.removeItem(LICENSE_INSTANCE_ID_KEY);
            localStorage.removeItem(LICENSE_TIER_NAME_KEY);
        }
    } catch (e) {
        console.error('Error saving license status:', e);
    }
}

/**
 * Save seat usage/limit from Lemon Squeezy for display in Settings.
 */
function saveLicenseSeats(uses, quantity) {
    try {
        var u = Number(uses);
        var q = Number(quantity);
        if (!isNaN(u) && !isNaN(q) && u >= 0 && q >= 0) {
            localStorage.setItem(LICENSE_SEAT_USES_KEY, String(Math.floor(u)));
            localStorage.setItem(LICENSE_SEAT_QUANTITY_KEY, String(Math.floor(q)));
        }
    } catch (e) {
        console.error('Error saving license seats:', e);
    }
}

/**
 * Get display string for seats: "current_uses / total_seats" or "—" when inactive / no data.
 */
function getLicenseSeatsDisplay() {
    try {
        if (!isLicenseVerified()) return '—';
        var uses = localStorage.getItem(LICENSE_SEAT_USES_KEY);
        var quantity = localStorage.getItem(LICENSE_SEAT_QUANTITY_KEY);
        if (uses != null && quantity != null && uses !== '' && quantity !== '') return uses + ' / ' + quantity;
        return '—';
    } catch (e) {
        return '—';
    }
}

/**
 * Get display string for license tier (from Lemon Squeezy meta.variant_name or stored value).
 */
function getLicenseTierDisplay() {
    try {
        if (!isLicenseVerified()) return '—';
        var name = localStorage.getItem(LICENSE_TIER_NAME_KEY);
        return (name && name.trim()) ? name.trim() : '—';
    } catch (e) {
        return '—';
    }
}

/**
 * Core validate + sync handler against Lemon Squeezy.
 * - logoutOnInvalid: when true, clears local state and forces license re-entry if remote says invalid/inactive/disabled/expired.
 * - done: optional callback always called at end.
 */
function validateAndSyncLicense(logoutOnInvalid, done) {
    if (!isLicenseVerified()) { if (done) done(); return; }
    var key = null;
    var instanceId = null;
    try {
        key = localStorage.getItem(LICENSE_KEY_FOR_RELEASE_KEY);
        instanceId = localStorage.getItem(LICENSE_INSTANCE_ID_KEY);
    } catch (e) { if (done) done(); return; }
    if (!key || key.trim().length < 8) { if (done) done(); return; }
    var body = 'license_key=' + encodeURIComponent(key.trim());
    if (instanceId && instanceId.trim()) body += '&instance_id=' + encodeURIComponent(instanceId.trim());
    var xhr = new XMLHttpRequest();
    xhr.open('POST', LEMON_SQUEEZY_VALIDATE_URL, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    xhr.setRequestHeader('Accept', 'application/json');
    xhr.timeout = 8000;
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status === 200) {
            try {
                var data = JSON.parse(xhr.responseText);
                if (data && data.license_key) {
                    var lk = data.license_key;
                    var status = (lk.status || '').toLowerCase();
                    var isValid = (data.valid === true) && (status === '' || status === 'active');
                    if (isValid) {
                        var usage = (lk.activation_usage != null) ? Number(lk.activation_usage) : 0;
                        var limit = (lk.activation_limit != null) ? Number(lk.activation_limit) : 1;
                        if (!isNaN(usage) && !isNaN(limit)) saveLicenseSeats(usage, limit);
                        var meta = data.meta || {};
                        var tierName = (meta.variant_name && meta.variant_name.trim()) ? meta.variant_name.trim()
                            : (meta.variant_id != null && LEMON_VARIANT_NAMES[meta.variant_id]) ? LEMON_VARIANT_NAMES[meta.variant_id] : null;
                        if (tierName) try { localStorage.setItem(LICENSE_TIER_NAME_KEY, tierName); } catch (e) {}
                    } else if (logoutOnInvalid) {
                        // Remote says this license or instance is no longer valid: force local logout
                        saveLicenseStatus(false);
                        resetLicenseFormToInitialState();
                        showOverlay();
                    }
                } else if (logoutOnInvalid) {
                    saveLicenseStatus(false);
                    resetLicenseFormToInitialState();
                    showOverlay();
                }
            } catch (e) {
                // If parsing fails and logoutOnInvalid, err on side of forcing re-login
                if (logoutOnInvalid) {
                    saveLicenseStatus(false);
                    resetLicenseFormToInitialState();
                    showOverlay();
                }
            }
        } else if (logoutOnInvalid && (xhr.status === 400 || xhr.status === 404 || xhr.status === 410 || xhr.status === 422)) {
            // Known “license invalid/not found” style codes – force logout
            saveLicenseStatus(false);
            resetLicenseFormToInitialState();
            showOverlay();
        }
        if (done) done();
    };
    xhr.ontimeout = function () { if (done) done(); };
    xhr.onerror = function () { if (done) done(); };
    xhr.send(body);
}

/**
 * Fetch current usage/limit and tier from Lemon Squeezy validate endpoint; save for UI.
 * Does NOT force logout on invalid; used for Settings seat refresh.
 */
function refreshSeatDataFromApi(done) {
    validateAndSyncLicense(false, done);
}

/**
 * Get display label for license status (Active/Pro or Inactive)
 */
function getLicenseStatusLabel() {
    return isLicenseVerified() ? 'Active' : 'Inactive';
}

/**
 * Get shortened license key for display (XXXX-XXXX-1234)
 */
function getLicenseDisplayKey() {
    try {
        if (!isLicenseVerified()) return '—';
        var last4 = localStorage.getItem(LICENSE_KEY_LAST4_KEY);
        return last4 ? 'XXXX-XXXX-' + last4 : '••••-••••-••••';
    } catch (e) {
        return '—';
    }
}

/**
 * Simple hash function for license key (not cryptographic, just for basic validation)
 */
function simpleHash(str) {
    var hash = 0;
    for (var i = 0; i < str.length; i++) {
        var char = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // Convert to 32bit integer
    }
    return Math.abs(hash).toString(16);
}

/**
 * Reset the license overlay form to its initial state (e.g. after deactivation or to recover from a stuck "Verifying..." state).
 */
function resetLicenseFormToInitialState() {
    resetButton();
    hideError();
    var keyInput = document.getElementById('license-key-input');
    if (keyInput) {
        keyInput.value = '';
        keyInput.disabled = false;
    }
}

/**
 * Deactivate license: call Lemon Squeezy deactivate, then clear stored license and show activation overlay.
 */
function deactivateLicense() {
    var key = null;
    var instanceId = null;
    try {
        key = localStorage.getItem(LICENSE_KEY_FOR_RELEASE_KEY);
        instanceId = localStorage.getItem(LICENSE_INSTANCE_ID_KEY);
    } catch (e) {}
    if (key && key.trim() && instanceId && instanceId.trim()) {
        var body = 'license_key=' + encodeURIComponent(key.trim()) + '&instance_id=' + encodeURIComponent(instanceId.trim());
        var xhr = new XMLHttpRequest();
        xhr.open('POST', LEMON_SQUEEZY_DEACTIVATE_URL, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.setRequestHeader('Accept', 'application/json');
        xhr.timeout = 8000;
        xhr.onreadystatechange = function () {
            if (xhr.readyState !== 4) return;
            finishDeactivate();
        };
        xhr.ontimeout = finishDeactivate;
        xhr.onerror = finishDeactivate;
        xhr.send(body);
    } else {
        finishDeactivate();
    }
    function finishDeactivate() {
        saveLicenseStatus(false);
        resetLicenseFormToInitialState();
        showOverlay();
    }
}

/**
 * Check if license is verified
 */
function isLicenseVerified() {
    try {
        return localStorage.getItem(LICENSE_STORAGE_KEY) === 'true';
    } catch (e) {
        return false;
    }
}

/**
 * Show license overlay
 */
function showOverlay() {
    var overlay = document.getElementById('license-overlay');
    if (overlay) {
        overlay.classList.remove('hidden');
        document.body.classList.add('license-overlay-active');
    }
}

/**
 * Hide license overlay
 */
function hideOverlay() {
    var overlay = document.getElementById('license-overlay');
    if (overlay) {
        overlay.classList.add('hidden');
        document.body.classList.remove('license-overlay-active');
    }
}

/**
 * Show error message
 */
function showError(message) {
    var errorMsg = document.getElementById('license-error');
    if (errorMsg) {
        errorMsg.textContent = message || 'Invalid license key. Please check and try again.';
        errorMsg.classList.add('show');
    }
}

/**
 * Hide error message
 */
function hideError() {
    var errorMsg = document.getElementById('license-error');
    if (errorMsg) {
        errorMsg.classList.remove('show');
    }
}

/**
 * Show success message (brief)
 */
function showSuccess() {
    // Could show a toast notification here
    console.log('License verified successfully');
}

/**
 * Reset button state
 */
function resetButton() {
    var btn = document.getElementById('activate-btn');
    var keyInput = document.getElementById('license-key-input');
    if (btn) {
        btn.textContent = 'Verify License';
        btn.disabled = false;
    }
    if (keyInput) {
        keyInput.disabled = false;
    }
}

/**
 * Format license key input (preserve existing dashes, allow long keys)
 * Handles keys with or without dashes, up to 64+ characters
 * Ensures proper trimming and sanitization
 */
function formatLicenseKey(input) {
    if (!input) return;
    
    // Get current value and trim whitespace from both ends
    var value = input.value.trim();
    
    // Remove all non-alphanumeric characters except dashes
    value = value.replace(/[^A-Z0-9-]/gi, '').toUpperCase();
    
    // If the key already has dashes, preserve them (user might have pasted formatted key)
    // Otherwise, format with dashes every 8 characters for display
    // For long keys, we'll preserve the user's input format
    if (value.indexOf('-') === -1 && value.length > 8) {
        // Only auto-format if no dashes exist and key is longer than 8 chars
        // Format with dashes every 8 characters (XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX)
        var formatted = '';
        for (var i = 0; i < value.length; i++) {
            if (i > 0 && i % 8 === 0) {
                formatted += '-';
            }
            formatted += value[i];
        }
        input.value = formatted;
            } else {
        // Preserve existing format (user pasted formatted key)
        input.value = value;
    }
}

/**
 * Handle paste event to trim and clean license key
 * Ensures proper sanitization of pasted content
 */
function handleLicenseKeyPaste(e) {
    // Allow the paste to happen first
    setTimeout(function() {
        var input = e.target;
        if (input && input.id === 'license-key-input') {
            // Trim whitespace from beginning and end (critical for copy-paste)
            var value = input.value.trim();
            
            // Remove any non-alphanumeric characters except dashes
            value = value.replace(/[^A-Z0-9-]/gi, '').toUpperCase();
            
            // Apply formatting
            input.value = value;
            
            // Trigger format function to handle dashes
            formatLicenseKey(input);
            
            // Hide any previous errors
            hideError();
        }
    }, 0);
}

// Initialize license system on page load
document.addEventListener("DOMContentLoaded", function() {
    // Check if already verified
    if (isLicenseVerified()) {
        hideOverlay();
        // On launch, silently validate with Lemon Squeezy to ensure remote + local are in sync.
        // If the license was deactivated/disabled/expired on the dashboard, this will force logout.
        validateAndSyncLicense(true, function () {});
    } else {
        showOverlay();
    }

    // Setup form submission
    var licenseForm = document.getElementById('license-form');
    if (licenseForm) {
        licenseForm.addEventListener('submit', function(e) {
            e.preventDefault();
            var keyInput = document.getElementById('license-key-input');
            if (keyInput) {
                verifyLicense(keyInput.value);
            }
        });
    }

    // Format license key input as user types
    var keyInput = document.getElementById('license-key-input');
    if (keyInput) {
        // Handle input events (typing)
        keyInput.addEventListener('input', function() {
            formatLicenseKey(this);
            hideError();
        });
        
        // Handle paste events (to trim and clean pasted content)
        keyInput.addEventListener('paste', handleLicenseKeyPaste);
        
        // Handle Enter key to submit
        keyInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                licenseForm.dispatchEvent(new Event('submit'));
            }
        });
        
        // Handle blur to trim any trailing spaces
        keyInput.addEventListener('blur', function() {
            this.value = this.value.trim();
        });
    }
    
});

/* global CSInterface */
/**
 * Helpers CEP Extension - Main Application Logic
 * Handles UI interactions, JSX communication, and plugin functionality
 */
(function () {
  try { window.__HELPERS_APP_BUNDLE_LOADED__ = true; } catch (e) {}
  const cs = new CSInterface();
  function onDomReady(fn) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', fn);
    } else {
      fn();
    }
  }

  // ============================================
  // JSX Loading & Communication System
  // ============================================

  // JSX Loading State
  let _jsxLoaded = false;
  let _jsxLoadAttempts = 0;
  const MAX_LOAD_ATTEMPTS = 3;

  function _escapeForJsString(p) {
    return String(p).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  }

  /**
   * Normalize path for passing to ExtendScript File(): use forward slashes.
   * getSystemPath() returns OS-native paths (backslashes on Windows); ExtendScript
   * File() accepts forward slashes on both macOS and Windows, so we normalize once.
   */
  function _normalizePathForExtendScript(path) {
    return String(path).replace(/\\/g, '/');
  }

  // Ping function to verify JS/JSX bridge - CRUCIAL: Check ping FIRST before any other logic
  function pingJSX(callback) {
    cs.evalScript('try { if ($.global && typeof $.global.helpersPing === "function") { $.global.helpersPing() } else { "BRIDGE_ERROR: helpersPing not found" } } catch(e) { "BRIDGE_ERROR: " + e.toString(); }', function(res) {
      if (callback) callback(String(res));
    });
  }

  // Load JSX files using SystemPath.EXTENSION (cross-platform; no hard-coded paths)
  function loadJSX(callback) {
    if (_jsxLoaded) {
      if (callback) callback(true);
      return;
    }

    const extRoot = cs.getSystemPath(SystemPath.EXTENSION);
    if (!extRoot) {
      setStatus('ERROR: Cannot get extension path');
      if (callback) callback(false);
      return;
    }

    // Build path with forward slashes so ExtendScript File() works on macOS and Windows
    const hostPath = _escapeForJsString(_normalizePathForExtendScript(extRoot + '/jsx/host.jsx'));

    // First, check if ping function exists (from host.jsx auto-load via manifest)
    pingJSX(function(pingResult) {
      if (pingResult && pingResult.indexOf('OK') === 0) {
        // Ping successful, check if all functions are loaded
        const checkScript = 'try {' +
          'if (typeof $.global === "undefined") { "NO_GLOBAL" } else {' +
            'var lg = (typeof $.global.helpersGenerateLogoGrid === "function");' +
            'var csf = (typeof $.global.helpersGenerateClearspace === "function");' +
            'var bg = (typeof $.global.helpersGenerateBaseGrid === "function");' +
            'var ping = (typeof $.global.helpersPing === "function");' +
            'if (lg && csf && bg && ping) { "OK" } else { "NEED_LOAD" }' +
          '}' +
          '} catch(e) { "ERROR: " + e.toString(); }';

        cs.evalScript(checkScript, function(res) {
          const result = String(res);
          if (result === 'OK') {
            _jsxLoaded = true;
            if (callback) callback(true);
            return;
          }
          // Functions not all loaded, try loading host.jsx
          loadHostJSX(hostPath, callback);
        });
      } else {
        // Ping failed, try loading host.jsx
        loadHostJSX(hostPath, callback);
      }
    });
  }

  // Load host.jsx using absolute path
  function loadHostJSX(hostPath, callback) {
    _jsxLoadAttempts++;
    if (_jsxLoadAttempts > MAX_LOAD_ATTEMPTS) {
      setStatus('ERROR: Failed to load JSX after ' + MAX_LOAD_ATTEMPTS + ' attempts');
      if (callback) callback(false);
      return;
    }

    const loadScript = 'try {' +
      'if (typeof $.global === "undefined") { $.global = {}; }' +
      'var hostFile = new File("' + hostPath + '");' +
      'if (hostFile.exists) {' +
        '$.evalFile(hostFile);' +
        '// Wait a moment for functions to attach' +
        'var lg = (typeof $.global.helpersGenerateLogoGrid === "function");' +
        'var csf = (typeof $.global.helpersGenerateClearspace === "function");' +
        'var bg = (typeof $.global.helpersGenerateBaseGrid === "function");' +
        'var ping = (typeof $.global.helpersPing === "function");' +
        'if (lg && csf && bg && ping) { "OK" } else { "PARTIAL: lg=" + lg + " csf=" + csf + " bg=" + bg + " ping=" + ping }' +
      '} else { "ERROR: host.jsx not found at " + "' + hostPath + '" }' +
      '} catch(e) { "ERROR: " + e.toString() + " (line: " + (e.line || "?") + ")"; }';

    cs.evalScript(loadScript, function(res2) {
      const result2 = String(res2);
      if (result2 === 'OK') {
        _jsxLoaded = true;
        if (callback) callback(true);
      } else {
        setStatus('JSX load failed: ' + result2);
        // Retry after a short delay
        setTimeout(function() {
          loadJSX(callback);
        }, 500);
      }
    });
  }

  // Ensure JSX is loaded before executing any function
  function ensureJsxLoaded(callback) {
    loadJSX(function(success) {
      if (success && callback) {
        callback();
      } else if (!success && callback) {
        callback(); // Still call callback to prevent UI lock
      }
    });
  }

  // Verify function exists before calling
  function verifyFunction(funcName, callback) {
    cs.evalScript('try { ($.global && typeof $.global.' + funcName + ' === "function") ? "OK" : "MISSING" } catch(e) { "ERROR: " + e.toString(); }', function(res) {
      if (callback) callback(String(res) === 'OK');
    });
  }

  const $ = (id) => document.getElementById(id);
  const qsa = (sel) => Array.prototype.slice.call(document.querySelectorAll(sel));

  // ============================================
  // Utility Functions
  // ============================================
  
  /**
   * Safely convert value to number with fallback
   */
  function safeNum(v, fallback) {
    const n = Number(v);
    return Number.isFinite(n) ? n : fallback;
  }

  /**
   * Sanitize status text for human-readable display: plain English, no technical/emoji characters.
   */
  function sanitizeStatus(msg) {
    if (msg == null) return '';
    var s = String(msg).trim();
    // Keep only printable ASCII (space through tilde) so the status bar never shows garbled chars
    s = s.replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim();
    return s || 'Ready';
  }

  /**
   * Update status message in UI (sanitized for plain-English display)
   */
  function setStatus(msg) {
    const el = $('status');
    if (el) el.textContent = sanitizeStatus(msg);
  }

  function evalJsx(script, cb) {
    try {
      ensureJsxLoaded(function () {
        cs.evalScript(script, function (result) {
          if (cb) cb(result);
        });
      });
    } catch (e) {
      setStatus('evalScript error: ' + (e && e.message ? e.message : String(e)));
      if (cb) cb('ERROR: ' + String(e));
    }
  }

  /**
   * Legacy hook kept for UI compatibility.
   * Grid rendering now goes through helpersGenerateLogoGrid from host.jsx only.
   */
  function runGridTypeScript(gridType) {
    const valid = gridType === 'straightLinesGrid' || gridType === 'diagonalGrid' || gridType === 'circleGrid';
    if (!valid) {
      setStatus('Unknown grid type: ' + gridType);
      return;
    }
  }

  // Initialize JSX on panel load - CRUCIAL: Ping first to verify bridge
  onDomReady(function() {
    // First, ping to verify bridge connection
    pingJSX(function(pingResult) {
      if (pingResult && pingResult.indexOf('OK') === 0) {
        // Ping successful, bridge is working
        setStatus('Bridge connected');
        // Now load JSX if needed
        loadJSX(function(success) {
          if (success) {
            setStatus('Ready');
          } else {
            setStatus('JSX load failed');
          }
        });
      } else {
        // Ping failed, try loading JSX first
        setStatus('Connecting bridge...');
        loadJSX(function(success) {
          if (success) {
            // Retry ping after loading
            pingJSX(function(res) {
              if (res && res.indexOf('OK') === 0) {
                setStatus('Ready');
              } else {
                setStatus('Bridge error: ' + (res || 'Unknown'));
              }
            });
          } else {
            setStatus('Bridge connection failed');
          }
        });
      }
    });
  });

  // ============================================
  // Preferences Management
  // ============================================
  
  /**
   * Save user preference to localStorage
   */
  function savePref(key, val) {
    try { localStorage.setItem('helpers.' + key, String(val)); } catch (e) {}
  }
  
  /**
   * Load user preference from localStorage
   */
  function loadPref(key, fallback) {
    try {
      const v = localStorage.getItem('helpers.' + key);
      return v === null ? fallback : v;
    } catch (e) {
      return fallback;
    }
  }

  // ============================================
  // Settings Management
  // ============================================
  
  // Global Style: default gray; color is persisted and editable in Settings (Logo Grid only)
  const DEFAULT_GLOBAL_COLOR = '#808080';
  const DEFAULT_GLOBAL_STROKE = 0.5;
  let globalElementColor = DEFAULT_GLOBAL_COLOR;
  const globalStrokeWeight = DEFAULT_GLOBAL_STROKE;
  // Clearspace always uses fixed gray (unaffected by Global Color in Settings)
  const CLEARSPACE_GRAY = { r: 128, g: 128, b: 128 };

  // Must match ExtensionBundleVersion in CSXS/manifest.xml for update checks to stay in sync with remote update.json
  const CURRENT_VERSION = '1.2.0';
  const UPDATE_CHECK_URL = 'https://helperstudio.lemonsqueezy.com/checkout';

  // ===============================
  // Update System
  // ===============================
  // Remote versioning file: https://raw.githubusercontent.com/Kobi7991/helpers-updates/refs/heads/main/update.json
  const UPDATE_URL = 'https://raw.githubusercontent.com/Kobi7991/helpers-updates/refs/heads/main/update.json';

  /**
   * Get remote version string from update.json. Accepts latestVersion or ExtensionBundleVersion for manifest alignment.
   */
  function _getRemoteVersion(data) {
    if (!data) return '';
    var v = (typeof data.latestVersion === 'string' && data.latestVersion) ? data.latestVersion
      : (typeof data.ExtensionBundleVersion === 'string' && data.ExtensionBundleVersion) ? data.ExtensionBundleVersion
      : '';
    return v.trim();
  }

  /**
   * Compare semantic versions (e.g. 1.2.10 vs 1.2.3). Returns true if remote > current.
   */
  function _isNewerVersion(remote, current) {
    try {
      var r = (remote || '').split('.').map(function (n) { return parseInt(n, 10) || 0; });
      var c = (current || '').split('.').map(function (n) { return parseInt(n, 10) || 0; });
      for (var i = 0; i < Math.max(r.length, c.length); i++) {
        var rn = r[i] || 0;
        var cn = c[i] || 0;
        if (rn > cn) return true;
        if (rn < cn) return false;
      }
      return false;
    } catch (e) {
      return false;
    }
  }

  /**
   * Fetch update.json with cache buster. Uses fetch; falls back to XMLHttpRequest. Fails silently.
   * Calls callback(null, data) on success, callback(err) or no call on failure.
   */
  function _fetchUpdateInfo(callback) {
    var url = UPDATE_URL;
    if (!url || typeof callback !== 'function') return;
    var cacheBustUrl = url + (url.indexOf('?') !== -1 ? '&' : '?') + 't=' + Date.now();

    function onSuccess(data) {
      try {
        callback(null, data);
      } catch (e) { /* fail silently */ }
    }
    function onFail() {
      try {
        callback(new Error('Update check failed'));
      } catch (e) { /* fail silently */ }
    }

    if (typeof window.fetch === 'function') {
      window.fetch(cacheBustUrl, { cache: 'no-store' })
        .then(function (res) {
          if (!res || res.status !== 200) { onFail(); return; }
          return res.text();
        })
        .then(function (text) {
          if (!text) { onFail(); return; }
          try {
            onSuccess(JSON.parse(text));
          } catch (e) {
            onFail();
          }
        })
        .catch(onFail);
      return;
    }

    var xhr = new XMLHttpRequest();
    xhr.open('GET', cacheBustUrl, true);
    xhr.timeout = 8000;
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;
      if (xhr.status !== 200) { onFail(); return; }
      try {
        onSuccess(JSON.parse(xhr.responseText));
      } catch (e) {
        onFail();
      }
    };
    xhr.onerror = onFail;
    xhr.ontimeout = onFail;
    try {
      xhr.send();
    } catch (e) {
      onFail();
    }
  }

  /** Pending update info when a newer version was detected: { version, downloadUrl } or null. */
  var _pendingUpdate = null;

  function _setPendingUpdate(data) {
    var latest = _getRemoteVersion(data);
    if (!latest) return;
    var downloadUrl = (data && data.downloadUrl && String(data.downloadUrl)) ? data.downloadUrl : UPDATE_CHECK_URL;
    _pendingUpdate = { version: latest, downloadUrl: downloadUrl };
  }

  function _clearPendingUpdate() {
    _pendingUpdate = null;
  }

  function _getPendingUpdate() {
    return _pendingUpdate;
  }

  /**
   * Show update dialog with message and clickable download link (opens in default browser).
   */
  function _showUpdateDialog(data) {
    var latest = _getRemoteVersion(data);
    if (!latest) return;
    var name = (data && data.pluginName && String(data.pluginName)) || 'helpers';
    var notes = (data.releaseNotes && String(data.releaseNotes)) ? String(data.releaseNotes) : '';
    var downloadUrl = (data.downloadUrl && String(data.downloadUrl)) ? data.downloadUrl : UPDATE_CHECK_URL;

    var overlay = $('updateDialogOverlay');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'updateDialogOverlay';
      overlay.className = 'update-dialog-overlay';
      overlay.setAttribute('aria-hidden', 'true');
      overlay.innerHTML = '<div class="update-dialog-box" role="dialog" aria-labelledby="updateDialogTitle" aria-modal="true">' +
        '<div class="update-dialog-header">' +
        '<h2 id="updateDialogTitle" class="update-dialog-title"></h2>' +
        '<button type="button" class="update-dialog-close" id="updateDialogClose" aria-label="Close">×</button>' +
        '</div>' +
        '<div class="update-dialog-body">' +
        '<p class="update-dialog-notes" id="updateDialogNotes"></p>' +
        '<a href="#" id="updateDialogLink" class="update-dialog-link" target="_blank">Download update</a>' +
        '</div></div>';
      document.body.appendChild(overlay);

      overlay.addEventListener('click', function (e) {
        if (e.target === overlay) _closeUpdateDialog();
      });
      var closeBtn = $('updateDialogClose');
      if (closeBtn) closeBtn.addEventListener('click', _closeUpdateDialog);
      document.addEventListener('keydown', function onKey(e) {
        if (e.key === 'Escape' && overlay.classList.contains('update-dialog-overlay--open')) {
          _closeUpdateDialog();
        }
      });
      var linkEl = $('updateDialogLink');
      if (linkEl) {
        linkEl.addEventListener('click', function (e) {
          e.preventDefault();
          try { cs.openURLInDefaultBrowser(downloadUrl); } catch (err) { window.open(downloadUrl, '_blank'); }
        });
      }
    }

    var titleEl = overlay.querySelector('#updateDialogTitle');
    var notesEl = overlay.querySelector('#updateDialogNotes');
    var linkEl = overlay.querySelector('#updateDialogLink');
    if (titleEl) titleEl.textContent = 'A new version (v' + latest + ') is available!';
    if (notesEl) { notesEl.textContent = notes; notesEl.style.display = notes ? 'block' : 'none'; }
    if (linkEl) {
      linkEl.href = downloadUrl;
      linkEl.textContent = 'Download update';
      linkEl.onclick = function (e) {
        e.preventDefault();
        try { cs.openURLInDefaultBrowser(downloadUrl); } catch (err) { window.open(downloadUrl, '_blank'); }
      };
    }
    overlay.classList.add('update-dialog-overlay--open');
    overlay.setAttribute('aria-hidden', 'false');
  }

  function _closeUpdateDialog() {
    var overlay = $('updateDialogOverlay');
    if (!overlay) return;
    overlay.classList.remove('update-dialog-overlay--open');
    overlay.setAttribute('aria-hidden', 'true');
  }

  /**
   * Ensure update badge element exists on settings button; return the element.
   */
  function _ensureUpdateBadge() {
    var btn = $('settingsBtn');
    if (!btn) return null;
    var badge = btn.querySelector('.update-badge');
    if (badge) return badge;
    badge = document.createElement('span');
    badge.className = 'update-badge';
    badge.setAttribute('aria-hidden', 'true');
    btn.appendChild(badge);
    return badge;
  }

  /**
   * Show or hide the blue update badge on the settings gear icon.
   */
  function setUpdateBadge(visible) {
    var badge = _ensureUpdateBadge();
    if (!badge) return;
    badge.classList.toggle('update-badge--visible', !!visible);
  }

  /**
   * Update the status text below the Check for Updates button. state: 'default' | 'current' | 'available'.
   */
  function _setUpdateStatusText(msg, state) {
    var el = $('settingsUpdateStatus');
    if (!el) return;
    el.textContent = msg || 'Current Version: v' + CURRENT_VERSION;
    el.className = 'settings-update-status';
    if (state === 'current') el.classList.add('settings-update-status--current');
    else if (state === 'available') el.classList.add('settings-update-status--available');
  }

  /**
   * Refresh the "Check for Updates" button: if a pending update exists, show "Download vX.Y.Z"; otherwise "Check for updates".
   */
  function _refreshUpdateButton() {
    var checkBtn = $('settingsCheckUpdates');
    if (!checkBtn) return;
    var pending = _getPendingUpdate();
    if (pending) {
      checkBtn.textContent = 'Download v' + pending.version;
      checkBtn.disabled = false;
    } else {
      checkBtn.textContent = 'Check for updates';
      checkBtn.disabled = false;
    }
  }

  /**
   * Automatic update check on startup. Fetches UPDATE_URL, compares remote version to CURRENT_VERSION; updates UI (no alerts).
   */
  function checkForUpdates() {
    _fetchUpdateInfo(function (err, data) {
      if (err || !data) return;
      var latest = _getRemoteVersion(data);
      if (!latest) return;
      if (_isNewerVersion(latest, CURRENT_VERSION)) {
        _setPendingUpdate(data);
        _showUpdateDialog(data);
        setUpdateBadge(true);
        _refreshUpdateButton();
        _setUpdateStatusText('New version available!', 'available');
      }
    });
  }

  /**
   * Manual update check (from "Check for Updates" button). Updates UI only (no alerts); button shows "Checking..." then reverts to "Check for updates" when current.
   */
  function runManualUpdateCheck() {
    var checkBtn = $('settingsCheckUpdates');
    if (checkBtn) {
      checkBtn.disabled = true;
      checkBtn.textContent = 'Checking...';
    }
    _fetchUpdateInfo(function (err, data) {
      if (err || !data) {
        if (checkBtn) { checkBtn.disabled = false; checkBtn.textContent = 'Check for updates'; }
        _refreshUpdateButton();
        return;
      }
      var latest = _getRemoteVersion(data);
      if (!latest) {
        if (checkBtn) { checkBtn.disabled = false; checkBtn.textContent = 'Check for updates'; }
        _refreshUpdateButton();
        return;
      }
      if (_isNewerVersion(latest, CURRENT_VERSION)) {
        _setPendingUpdate(data);
        _showUpdateDialog(data);
        setUpdateBadge(true);
        _refreshUpdateButton();
        _setUpdateStatusText('New version available!', 'available');
        if (checkBtn) checkBtn.disabled = false;
      } else {
        _clearPendingUpdate();
        setUpdateBadge(false);
        _setUpdateStatusText('You are up to date (v' + CURRENT_VERSION + ')', 'current');
        if (checkBtn) checkBtn.disabled = false;
        _refreshUpdateButton();
        setTimeout(function () {
          if (checkBtn) checkBtn.textContent = 'Check for updates';
        }, 1000);
      }
    });
  }

  function refreshSettingsStatus() {
    const statusEl = $('settingsLicenseStatus');
    const tierEl = $('settingsLicenseTier');
    const seatsEl = $('settingsLicenseSeats');
    const keyEl = $('settingsLicenseKey');
    const deactivateBtn = $('settingsDeactivateLicense');
    const label = getLicenseStatusLabel();
    if (statusEl) {
      statusEl.textContent = label;
      statusEl.classList.toggle('license-active', label === 'Active');
    }
    if (tierEl) tierEl.textContent = getLicenseTierDisplay();
    if (seatsEl) seatsEl.textContent = getLicenseSeatsDisplay();
    if (isLicenseVerified()) {
      refreshSeatDataFromApi(function () {
        if (tierEl) tierEl.textContent = getLicenseTierDisplay();
        if (seatsEl) seatsEl.textContent = getLicenseSeatsDisplay();
      });
    }
    if (keyEl) keyEl.textContent = getLicenseDisplayKey();
    if (deactivateBtn) {
      deactivateBtn.style.display = isLicenseVerified() ? '' : 'none';
    }
    var globalColorPickerEl = $('settingsGlobalColorPicker');
    var globalColorHexEl = $('settingsGlobalColorHex');
    if (globalColorPickerEl) globalColorPickerEl.value = globalElementColor;
    if (globalColorHexEl) globalColorHexEl.value = globalElementColor;
    var pending = _getPendingUpdate();
    if (pending) {
      _setUpdateStatusText('New version available!', 'available');
    } else {
      _setUpdateStatusText('Current Version: v' + CURRENT_VERSION, 'default');
    }
    _refreshUpdateButton();
  }

  function initializeSettings() {
    const settingsBtn = $('settingsBtn');
    const settingsOverlay = $('settingsOverlay');
    const settingsClose = $('settingsClose');
    const checkUpdatesBtn = $('settingsCheckUpdates');

    if (!settingsBtn || !settingsOverlay) return;

    settingsBtn.addEventListener('click', function() {
      refreshSettingsStatus();
      settingsOverlay.classList.add('active');
    });
    if (settingsClose) settingsClose.addEventListener('click', function() {
      settingsOverlay.classList.remove('active');
    });
    settingsOverlay.addEventListener('click', function(e) {
      if (e.target === settingsOverlay) settingsOverlay.classList.remove('active');
    });
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape' && settingsOverlay.classList.contains('active')) {
        settingsOverlay.classList.remove('active');
      }
    });

    if (checkUpdatesBtn) {
      checkUpdatesBtn.addEventListener('click', function () {
        var pending = _getPendingUpdate();
        if (pending) {
          try { cs.openURLInDefaultBrowser(pending.downloadUrl); } catch (e) { window.open(pending.downloadUrl, '_blank'); }
        } else {
          runManualUpdateCheck();
        }
      });
    }

    var deactivateLicenseBtn = $('settingsDeactivateLicense');
    if (deactivateLicenseBtn) {
      deactivateLicenseBtn.addEventListener('click', function () {
        deactivateLicense();
        refreshSettingsStatus();
        settingsOverlay.classList.remove('active');
      });
    }

    // Global Color for Logo Grid: load saved value and wire hex input + color picker preview
    var savedHex = loadPref('globalColor', DEFAULT_GLOBAL_COLOR);
    globalElementColor = normalizeHex(savedHex);
    var globalColorPicker = $('settingsGlobalColorPicker');
    var globalColorHex = $('settingsGlobalColorHex');
    if (globalColorPicker && globalColorHex) {
      globalColorPicker.value = globalElementColor;
      globalColorHex.value = globalElementColor;
      function applyGlobalColor(hex) {
        hex = normalizeHex(hex);
        globalElementColor = hex;
        savePref('globalColor', hex);
        globalColorPicker.value = hex;
        globalColorHex.value = hex;
        applyGlobalStyleToExistingElementsOnly();
      }
      globalColorPicker.addEventListener('input', function () { applyGlobalColor(globalColorPicker.value); });
      globalColorPicker.addEventListener('change', function () { applyGlobalColor(globalColorPicker.value); });
      globalColorHex.addEventListener('input', function () {
        var raw = globalColorHex.value.trim();
        if (/^#?[a-fA-F0-9]{6}$/.test(raw) || /^[a-fA-F0-9]{3}$/.test(raw)) {
          applyGlobalColor(raw.indexOf('#') === 0 ? raw : '#' + raw);
        }
      });
      globalColorHex.addEventListener('change', function () {
        var raw = globalColorHex.value.trim();
        if (raw) applyGlobalColor(raw.indexOf('#') === 0 ? raw : '#' + raw);
      });
    }
  }

  // Color Picker - Custom overlay. Returns normalized #rrggbb for Illustrator RGBColor compatibility.
  function normalizeHex(hex) {
    if (!hex || typeof hex !== 'string') return DEFAULT_GLOBAL_COLOR;
    const s = hex.replace(/^#/, '').trim().toLowerCase();
    if (s.length === 6 && /^[a-f0-9]{6}$/.test(s)) return '#' + s;
    if (s.length === 3 && /^[a-f0-9]{3}$/.test(s)) return '#' + s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
    return DEFAULT_GLOBAL_COLOR;
  }
  function hexToHsl(hex) {
    const r = parseInt(hex.slice(1, 3), 16) / 255;
    const g = parseInt(hex.slice(3, 5), 16) / 255;
    const b = parseInt(hex.slice(5, 7), 16) / 255;
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h = 0;
    let s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      switch (max) {
        case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
        case g: h = ((b - r) / d + 2) / 6; break;
        default: h = ((r - g) / d + 4) / 6; break;
      }
    }
    return { h: h * 360, s: s * 100, l: l * 100 };
  }
  function hslToHex(h, s, l) {
    l /= 100;
    const a = (s / 100) * Math.min(l, 1 - l);
    const f = function(n) {
      const k = (n + h / 30) % 12;
      const c = l - a * Math.max(Math.min(k - 3, 9 - k, 1), -1);
      return Math.round(255 * Math.max(0, Math.min(1, c)));
    };
    return '#' + [f(0), f(8), f(4)].map(function(x) { return x.toString(16).padStart(2, '0'); }).join('');
  }

  function initializeColorPickerModal() {
    const overlay = $('colorPickerOverlay');
    const container = $('colorPickerContainer');
    const backdrop = $('colorPickerBackdrop');
    const satLight = $('colorPickerSatLight');
    const cursor = $('colorPickerCursor');
    const hueInput = $('colorPickerHue');
    const hexInput = $('colorPickerHex');
    const doneBtn = $('colorPickerDone');
    if (!overlay || !container || !satLight || !cursor || !hueInput || !hexInput || !doneBtn) return;

    let state = { h: 210, s: 100, l: 50 };
    let targetColorInput = null;
    let targetTextInput = null;

    function updateSatLightBg() {
      const hueColor = hslToHex(state.h, 100, 50);
      satLight.style.background = 'linear-gradient(to top, #000, transparent), linear-gradient(to right, #fff, ' + hueColor + ')';
    }
    function moveCursor() {
      const x = (state.s / 100) * 100;
      const y = 100 - (state.l / 100) * 100;
      cursor.style.left = x + '%';
      cursor.style.top = y + '%';
    }
    function stateToHex() {
      return hslToHex(state.h, state.s, state.l);
    }
    function syncUI() {
      updateSatLightBg();
      moveCursor();
      hueInput.value = Math.round(state.h);
      const hex = normalizeHex(stateToHex());
      hexInput.value = hex;
      if (targetColorInput) {
        targetColorInput.value = hex;
        if (targetTextInput) targetTextInput.value = hex;
      }
    }

    function applyAndClose() {
      if (!targetColorInput) return;
      const hex = normalizeHex(stateToHex());
      targetColorInput.value = hex;
      if (targetTextInput) targetTextInput.value = hex;
      targetColorInput.dispatchEvent(new Event('change', { bubbles: true }));
      targetColorInput.dispatchEvent(new Event('input', { bubbles: true }));
      overlay.classList.add('hidden');
      overlay.setAttribute('aria-hidden', 'true');
      targetColorInput = null;
      targetTextInput = null;
    }

    function openPicker(colorInput) {
      const hex = normalizeHex((colorInput.value || '#2f8cff').trim());
      const hsl = hexToHsl(hex);
      state = { h: hsl.h, s: hsl.s, l: hsl.l };
      targetColorInput = colorInput;
      const textId = colorInput.getAttribute('data-color-text');
      targetTextInput = textId ? $(textId) : null;
      syncUI();
      overlay.classList.remove('hidden');
      overlay.setAttribute('aria-hidden', 'false');
    }

    hueInput.addEventListener('input', function() {
      state.h = Number(hueInput.value);
      syncUI();
    });
    satLight.addEventListener('mousedown', function(e) {
      const rect = satLight.getBoundingClientRect();
      const w = rect.width;
      const h = rect.height;
      const x = Math.max(0, Math.min(1, (e.clientX - rect.left) / w));
      const y = Math.max(0, Math.min(1, (e.clientY - rect.top) / h));
      state.s = x * 100;
      state.l = (1 - y) * 100;
      syncUI();
      function move(ev) {
        const xx = Math.max(0, Math.min(1, (ev.clientX - rect.left) / w));
        const yy = Math.max(0, Math.min(1, (ev.clientY - rect.top) / h));
        state.s = xx * 100;
        state.l = (1 - yy) * 100;
        syncUI();
      }
      function up() {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
      }
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
    });
    hexInput.addEventListener('input', function() {
      const raw = hexInput.value.trim();
      if (/^#?[a-fA-F0-9]{6}$/.test(raw) || /^[a-fA-F0-9]{3}$/.test(raw)) {
        const hex = normalizeHex(raw.startsWith('#') ? raw : '#' + raw);
        const hsl = hexToHsl(hex);
        state = { h: hsl.h, s: hsl.s, l: hsl.l };
        hexInput.value = hex;
        updateSatLightBg();
        moveCursor();
        hueInput.value = Math.round(state.h);
        if (targetColorInput) {
          targetColorInput.value = hex;
          if (targetTextInput) targetTextInput.value = hex;
        }
      }
    });
    doneBtn.addEventListener('click', applyAndClose);
    backdrop.addEventListener('click', applyAndClose);
    container.addEventListener('click', function(e) { e.stopPropagation(); });

    qsa('.settings-color-input').forEach(function(colorInput) {
      colorInput.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        openPicker(colorInput);
      });
    });
  }

  function hexToRgb(hex) {
    const h = normalizeHex(hex);
    const result = /^#([a-f0-9]{2})([a-f0-9]{2})([a-f0-9]{2})$/i.exec(h);
    return result ? {
      r: parseInt(result[1], 16),
      g: parseInt(result[2], 16),
      b: parseInt(result[3], 16)
    } : { r: 128, g: 128, b: 128 };
  }

  function getGlobalColor() {
    return hexToRgb(globalElementColor);
  }

  function getGlobalStrokeWeight() {
    return Number(globalStrokeWeight) || DEFAULT_GLOBAL_STROKE;
  }

  // Initialize settings on DOM ready
  onDomReady(function() {
    initializeSettings();
  });

  function setSegActive(segEl, value, attrName) {
    if (!segEl) return;
    const btns = Array.from(segEl.querySelectorAll('.segBtn'));
    btns.forEach(b => {
      const v = b.getAttribute(attrName);
      b.classList.toggle('active', v === value);
      b.setAttribute('aria-selected', v === value ? 'true' : 'false');
    });
  }

  // ============================================
  // View Management
  // ============================================
  
  const views = {
    logo: $('view-logo'),
    clear: $('view-clear'),
    base: $('view-base')
  };
  const tabBtns = qsa('#mainTabs .segBtn');
  let currentView = loadPref('view', 'logo');

  function updatePrimaryLabel() {
    const btn = $('btnPrimary');
    if (!btn) return;
    if (currentView === 'logo') btn.textContent = 'Generate';
    if (currentView === 'clear') btn.textContent = 'Generate';
    if (currentView === 'base') btn.textContent = 'Generate';
  }

  function showView(name) {
    currentView = name;
    savePref('view', name);
    tabBtns.forEach(b => {
      const on = b.dataset.view === name;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    Object.keys(views).forEach(k => {
      const el = views[k];
      if (el) el.classList.toggle('hidden', k !== name);
    });
    updatePrimaryLabel();
    
    if (name === 'clear') {
      updateAllComponentDisplays();
    }
    
    // Update Live Preview indicator visibility
    updateLivePreviewIndicator();
  }
  
  function updateLivePreviewIndicator() {
    const indicator = $('csLivePreviewIndicator');
    if (!indicator) return;
    
    // Show indicator only if live preview is enabled and objects have been generated for current view
    // Clear Space does NOT use live preview, so exclude it from the indicator
    const shouldShow = livePreviewEnabled && 
      ((currentView === 'logo' && objectsGenerated.logo) ||
       (currentView === 'base' && objectsGenerated.base));
    
    if (shouldShow) {
      indicator.classList.remove('hidden');
    } else {
      indicator.classList.add('hidden');
    }
  }

  tabBtns.forEach(b => b.addEventListener('click', () => showView(b.dataset.view)));

  // ============================================
  // UI Component Initialization
  // ============================================
  
  /**
   * Initialize accordion components with accordion behavior (only one open at a time)
   */
  function initAccordions() {
    // Handle new accordion card headers
    const accordionCardHeaders = qsa('.accordionCardHeader');
    const accordionCards = qsa('.accordionCard');
    
    accordionCardHeaders.forEach(header => {
      header.addEventListener('click', function(e) {
        // Don't toggle accordion if clicking the icon area (that's for tile toggle)
        if (e.target.closest('.accordionCardIcon')) {
          return;
        }
        
        e.preventDefault();
        e.stopPropagation();
        
        const isExpanded = this.getAttribute('aria-expanded') === 'true';
        const contentId = this.getAttribute('aria-controls');
        const card = this.closest('.accordionCard');
        const view = this.closest('.view');
        
        if (!card || !view) return;
        
        // If clicking to expand, close all other cards in the same view
        if (!isExpanded) {
          const allCardsInView = view.querySelectorAll('.accordionCard');
          allCardsInView.forEach(otherCard => {
            if (otherCard !== card) {
              const otherHeader = otherCard.querySelector('.accordionCardHeader');
              const otherContent = otherCard.querySelector('.accordionCardContent');
              if (otherHeader && otherContent) {
                otherHeader.setAttribute('aria-expanded', 'false');
              }
            }
          });
        }
        
        // Toggle current card
        this.setAttribute('aria-expanded', !isExpanded);
        
        // Save preference
        const moduleId = card.getAttribute('data-module');
        if (moduleId) {
          savePref('accordion.' + moduleId, !isExpanded ? '1' : '0');
        } else if (contentId) {
          savePref('accordion.' + contentId, !isExpanded ? '1' : '0');
        }
      });
      
      // Restore saved state (only first card open by default)
      const card = header.closest('.accordionCard');
      const view = header.closest('.view');
      if (!card || !view) return;
      
      const moduleId = card.getAttribute('data-module');
      const contentId = header.getAttribute('aria-controls');
      
      // Check if this is the first card in the view
      const allCardsInView = Array.from(view.querySelectorAll('.accordionCard'));
      const isFirstCard = allCardsInView.indexOf(card) === 0;
      
      let shouldBeOpen = false;
      if (moduleId) {
        const saved = loadPref('accordion.' + moduleId, isFirstCard ? '1' : '0');
        shouldBeOpen = saved === '1';
      } else if (contentId) {
        const saved = loadPref('accordion.' + contentId, isFirstCard ? '1' : '0');
        shouldBeOpen = saved === '1';
      }
      
      header.setAttribute('aria-expanded', shouldBeOpen ? 'true' : 'false');
      
      // If this card should be open, ensure others in the same view are closed
      if (shouldBeOpen) {
        allCardsInView.forEach(otherCard => {
          if (otherCard !== card) {
            const otherHeader = otherCard.querySelector('.accordionCardHeader');
            if (otherHeader) {
              otherHeader.setAttribute('aria-expanded', 'false');
            }
          }
        });
      }
    });
    
    // Handle legacy accordion headers (for Base Grid old structure if any)
    const legacyAccordionHeaders = qsa('.accordionHeader:not(.accordionCardHeader)');
    legacyAccordionHeaders.forEach(header => {
      header.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        const isExpanded = this.getAttribute('aria-expanded') === 'true';
        const contentId = this.getAttribute('aria-controls');
        
        // Toggle state
        this.setAttribute('aria-expanded', !isExpanded);
        
        // Save preference
        if (contentId) {
          savePref('accordion.' + contentId, !isExpanded ? '1' : '0');
        }
      });
      
      // Restore saved state
      const contentId = header.getAttribute('aria-controls');
      if (contentId) {
        const saved = loadPref('accordion.' + contentId, '1');
        header.setAttribute('aria-expanded', saved === '1' ? 'true' : 'false');
      }
    });
  }
  
  initAccordions();

  // Segmented Controls
  const gridDirSeg = $('gridDirSeg');

  function initSegment(segEl, key, hiddenInputId, attrName, fallback) {
    if (!segEl) return;
    const hidden = $(hiddenInputId);
    const saved = loadPref(key, fallback);
    if (hidden) hidden.value = saved;
    setSegActive(segEl, saved, attrName);

    segEl.addEventListener('click', (e) => {
      const btn = e.target.closest('.segBtn');
      if (!btn) return;
      const v = btn.getAttribute(attrName);
      if (!v) return;
      if (hidden) hidden.value = v;
      setSegActive(segEl, v, attrName);
      savePref(key, v);
    });
  }

  initSegment(gridDirSeg, 'gridDir', 'lgGridDir', 'data-dir', 'both');
  // Old clearspace mode segmented removed - replaced with Akrivi X Mode radio buttons

  // Gridlines direction controls were removed from UI (always uses hidden #lgGridDir = both).

  // Logo Grid Tiles (multi-select)
  const logoTiles = {
    anchors: $('tileAnchors'),
    handles: $('tileHandles'),
    gridlines: $('tileGridlines'),
    outlines: $('tileOutlines')
  };

  const logoCtrlBinds = qsa('[data-binds]');

  function tileOn(tileEl) {
    return tileEl && tileEl.classList.contains('on');
  }

  function setTile(tileEl, on) {
    if (!tileEl) return;
    const isOn = !!on;
    tileEl.classList.toggle('on', isOn);
    tileEl.setAttribute('aria-pressed', isOn ? 'true' : 'false');

    // Icon color/state swap (SVG via <img> src)
    // Support both old structure (tileIcon) and new structure (accordionCardIconImg)
    const icon = tileEl.querySelector('img.tileIcon') || tileEl.querySelector('img.accordionCardIconImg') || tileEl.querySelector('.accordionCardIcon img');
    if (icon) {
      const offSrc = icon.getAttribute('data-icon-off') || icon.getAttribute('src');
      const onSrc  = icon.getAttribute('data-icon-on')  || offSrc;
      icon.setAttribute('src', isOn ? onSrc : offSrc);
    }
  }
  function refreshBoundControls() {
    logoCtrlBinds.forEach(el => {
      const mod = el.getAttribute('data-binds');
      const isOn = tileOn(logoTiles[mod]);
      el.classList.toggle('disabled', !isOn);
      // disable inputs inside
      const inputs = Array.from(el.querySelectorAll('input,select,button'));
      inputs.forEach(i => {
        // allow segmented to remain clickable if enabled
        // Skip shape buttons - they should remain clickable
        if (!i.classList.contains('shapeBtn')) {
          i.disabled = !isOn;
        }
      });
    });
  }

  function initLogoTile(mod, prefKey) {
    const t = logoTiles[mod];
    if (!t) return;
    const saved = loadPref(prefKey, '1');
    setTile(t, saved !== '0');
    
    // For new accordion card structure, handle tile toggle on icon click
    if (t.classList.contains('accordionCardHeader')) {
      const iconArea = t.querySelector('.accordionCardIcon');
      if (iconArea) {
        iconArea.addEventListener('click', (e) => {
          e.stopPropagation(); // Prevent accordion toggle
          const on = !tileOn(t);
          setTile(t, on);
          savePref(prefKey, on ? '1' : '0');
          refreshBoundControls();
        });
        iconArea.style.cursor = 'pointer';
      }
    } else {
      // Legacy structure - entire tile toggles
      t.addEventListener('click', () => {
        const on = !tileOn(t);
        setTile(t, on);
        savePref(prefKey, on ? '1' : '0');
        refreshBoundControls();
      });
    }
  }

  initLogoTile('anchors', 'lg.anchors');
  initLogoTile('handles', 'lg.handles');
  initLogoTile('gridlines', 'lg.gridlines');
  initLogoTile('outlines', 'lg.outlines');

  // restore numeric prefs (Logo Grid strokes follow Outline pattern)
  const prefInputs = [
    ['lgAnchorSize', 'lg.anchorSize', '3'],
    ['lgAnchorStroke', 'lg.anchorStroke', '0.5'],
    ['lgHandleStroke', 'lg.handleStroke', '0.5'],
    ['lgOutlineStroke', 'lg.outlineStroke', '0.5'],
    ['lgGridlinesStroke', 'lg.gridlinesStroke', '0.5'],
    ['gridSize', 'base.size', '40'],
    ['strokeW', 'base.strokeW', '0.5']
  ];
  prefInputs.forEach(([id, key, fb]) => {
    const el = $(id);
    if (!el) return;
    el.value = loadPref(key, fb);
    el.addEventListener('change', () => savePref(key, el.value));
    el.addEventListener('input', () => savePref(key, el.value));
  });

  // Fill toggles for Anchors and Handles
  const anchorFillEl = $('lgAnchorFill');
  if (anchorFillEl) {
    anchorFillEl.checked = loadPref('lg.anchorFill', '1') !== '0';
    anchorFillEl.addEventListener('change', () => savePref('lg.anchorFill', anchorFillEl.checked ? '1' : '0'));
  }
  const handleFillEl = $('lgHandleFill');
  if (handleFillEl) {
    handleFillEl.checked = loadPref('lg.handleFill', '0') !== '0';
    handleFillEl.addEventListener('change', () => savePref('lg.handleFill', handleFillEl.checked ? '1' : '0'));
  }

  /**
   * Custom stepper UI: wrap numeric inputs with dark-theme increment/decrement arrows.
   * Replaces default HTML number spinners with professional, matching controls.
   */
  function initSteppers() {
    const inputs = document.querySelectorAll('input.accordionInput[type="number"]');
    const chevronUp = '<svg viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M6 4l3 4H3l3-4z" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    const chevronDown = '<svg viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M6 8l3-4H3l3 4z" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    inputs.forEach(function (input) {
      if (input.closest('.stepper')) return;
      var min = parseFloat(input.getAttribute('min'));
      var max = parseFloat(input.getAttribute('max'));
      var step = parseFloat(input.getAttribute('step')) || 1;
      if (isNaN(min)) min = -Infinity;
      if (isNaN(max)) max = Infinity;
      if (isNaN(step) || step <= 0) step = 1;
      input.classList.add('stepper-input');
      var wrapper = document.createElement('div');
      wrapper.className = 'stepper';
      input.parentNode.insertBefore(wrapper, input);
      wrapper.appendChild(input);
      var actions = document.createElement('div');
      actions.className = 'stepper-actions';
      var btnInc = document.createElement('button');
      btnInc.type = 'button';
      btnInc.className = 'stepper-btn stepper-btn-inc';
      btnInc.setAttribute('aria-label', 'Increase value');
      btnInc.innerHTML = chevronUp;
      var btnDec = document.createElement('button');
      btnDec.type = 'button';
      btnDec.className = 'stepper-btn stepper-btn-dec';
      btnDec.setAttribute('aria-label', 'Decrease value');
      btnDec.innerHTML = chevronDown;
      actions.appendChild(btnInc);
      actions.appendChild(btnDec);
      wrapper.appendChild(actions);
      function decimalsForStep(s) {
        if (s >= 1) return 0;
        var d = 0, x = s;
        while (x < 1 && d < 10) { x *= 10; d++; }
        return d;
      }
      function roundToDecimals(val, decimals) {
        if (decimals <= 0) return Math.round(val);
        var f = Math.pow(10, decimals);
        return Math.round(val * f) / f;
      }
      function clamp(val) {
        if (val < min) return min;
        if (val > max) return max;
        return val;
      }
      function formatStepperValue(val) {
        var clamped = clamp(val);
        var decimals = decimalsForStep(step);
        var rounded = roundToDecimals(clamped, decimals);
        return decimals <= 0 ? String(Math.round(rounded)) : String(rounded);
      }
      function updateButtons() {
        var v = parseFloat(input.value);
        if (isNaN(v)) return;
        btnInc.disabled = v >= max;
        btnDec.disabled = v <= min;
      }
      function stepUp() {
        var v = parseFloat(input.value);
        if (isNaN(v)) v = min;
        input.value = formatStepperValue(v + step);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        updateButtons();
      }
      function stepDown() {
        var v = parseFloat(input.value);
        if (isNaN(v)) v = max;
        input.value = formatStepperValue(v - step);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        updateButtons();
      }
      btnInc.addEventListener('click', stepUp);
      btnDec.addEventListener('click', stepDown);
      input.addEventListener('input', updateButtons);
      input.addEventListener('change', updateButtons);
      updateButtons();
    });
  }
  initSteppers();

  // Mirror controls inside module cards (so it matches the Figma layout)
  function mirrorNumeric(prefKey, primaryId, otherIds) {
    const primary = $(primaryId);
    const others = otherIds.map(id => $(id)).filter(Boolean);
    if (!primary || others.length === 0) return;

    function syncFromPrimary() {
      others.forEach(o => { o.value = primary.value; });
    }

    function syncFromAny(src) {
      primary.value = src.value;
      others.forEach(o => { if (o !== src) o.value = src.value; });
      savePref(prefKey, primary.value);
    }

    // initial
    syncFromPrimary();
    // primary edits
    primary.addEventListener('input', syncFromPrimary);
    primary.addEventListener('change', syncFromPrimary);
    // secondary edits
    others.forEach(o => {
      o.addEventListener('input', () => syncFromAny(o));
      o.addEventListener('change', () => syncFromAny(o));
    });
  }

  // NOTE: Anchor size and Handle size are now COMPLETELY INDEPENDENT
  // They no longer mirror each other - each has its own separate control
  // This ensures changing handle size doesn't affect anchor size and vice versa

  // ============================================
  // Shape Selector Initialization (Icon-based)
  // ============================================
  
  /**
   * Initialize shape selector with icon buttons
   * Each feature (Anchors/Handles) has its own independent selector
   */
  function initShapeSelector(selectorId, hiddenInputId, prefKey, defaultValue, logoGridFeature) {
    const selector = $(selectorId);
    const hiddenInput = $(hiddenInputId);
    if (!selector || !hiddenInput) return;
    
    // Load saved preference
    const savedShape = loadPref(prefKey, defaultValue);
    hiddenInput.value = savedShape;
    
    // Set initial active state
    const buttons = selector.querySelectorAll('.shapeBtn');
    buttons.forEach(btn => {
      const shape = btn.getAttribute('data-shape');
      const isActive = shape === savedShape;
      btn.classList.toggle('active', isActive);
      btn.setAttribute('aria-pressed', isActive ? 'true' : 'false');
    });
    
    // Add click handlers - each selector is completely independent; only update its feature
    buttons.forEach(btn => {
      btn.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        const shape = this.getAttribute('data-shape');
        
        // Update hidden input
        hiddenInput.value = shape;
        
        // Update active state
        buttons.forEach(b => {
          const isActive = b === this;
          b.classList.toggle('active', isActive);
          b.setAttribute('aria-pressed', isActive ? 'true' : 'false');
        });
        
        // Save preference
        savePref(prefKey, shape);
        
        // Run live update ONLY for this feature (anchors or handles) - no side effects on others
        if (logoGridFeature) runLogoGridLiveUpdate(logoGridFeature);
      });
    });
  }
  
  // Grid Type Selector (icon-based, multi-select toggle behavior)
  function initGridTypeSelector() {
    const selector = $('lgGridTypeSelector');
    const hiddenInput = $('lgGridType');
    if (!selector || !hiddenInput) return;
    
    // Load saved preference or use default (comma-separated string for multi-select)
    const savedGridTypes = loadPref('lg.gridType', 'straightLinesGrid');
    const selectedTypes = savedGridTypes.split(',').map(s => s.trim()).filter(s => s);
    
    // Ensure at least one type is selected (default to straightLinesGrid if none)
    const finalSelectedTypes = selectedTypes.length > 0 ? selectedTypes : ['straightLinesGrid'];
    const gridTypesString = finalSelectedTypes.join(',');
    hiddenInput.value = gridTypesString;
    
    // Set initial active state
    const buttons = selector.querySelectorAll('.shapeBtn');
    buttons.forEach(btn => {
      const gridType = btn.getAttribute('data-grid-type');
      const isActive = finalSelectedTypes.includes(gridType);
      btn.classList.toggle('active', isActive);
      btn.setAttribute('aria-pressed', isActive ? 'true' : 'false');
    });
    
    // Add click handlers - toggle behavior (multi-select)
    buttons.forEach(btn => {
      btn.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        const gridType = this.getAttribute('data-grid-type');
        const isCurrentlyActive = this.classList.contains('active');
        
        // Toggle this button's state FIRST
        const newActiveState = !isCurrentlyActive;
        this.classList.toggle('active', newActiveState);
        this.setAttribute('aria-pressed', newActiveState ? 'true' : 'false');
        
        // NOW collect all active grid types AFTER the toggle
        // Use a fresh query to ensure we get the current state
        const activeTypes = [];
        const allButtons = selector.querySelectorAll('.shapeBtn');
        allButtons.forEach(b => {
          if (b.classList.contains('active')) {
            const type = b.getAttribute('data-grid-type');
            if (type) activeTypes.push(type);
          }
        });
        
        // Ensure at least one type is selected (prevent empty selection)
        let finalActiveTypes = activeTypes;
        if (activeTypes.length === 0) {
          // No buttons active - activate default
          const defaultBtn = selector.querySelector('[data-grid-type="straightLinesGrid"]');
          if (defaultBtn) {
            defaultBtn.classList.add('active');
            defaultBtn.setAttribute('aria-pressed', 'true');
            finalActiveTypes = ['straightLinesGrid'];
          } else {
            finalActiveTypes = ['straightLinesGrid'];
          }
        }
        
        const gridTypesString = finalActiveTypes.join(',');
        hiddenInput.value = gridTypesString;
        
        // Save preference
        savePref('lg.gridType', gridTypesString);
        
        // Run live update only for gridlines - no side effects on anchors/handles/outline
        runLogoGridLiveUpdate('gridlines');
        
        // Keep compatibility hook for grid type icon clicks (host.jsx now owns all grid logic).
        runGridTypeScript(gridType);
      });
    });
  }
  
  // Initialize Anchors shape selector (independent) - live update only touches Anchors
  initShapeSelector('lgAnchorShapeSelector', 'lgAnchorShape', 'lg.anchorShape', 'square', 'anchors');
  
  // Initialize Handles shape selector (independent) - live update only touches Handles
  initShapeSelector('lgHandleShapeSelector', 'lgHandleShape', 'lg.handleShape', 'circle', 'handles');
  
  // Initialize Grid type selector (icon-based)
  initGridTypeSelector();

  // checkboxes
  const useSelection = $('useSelection');
  if (useSelection) {
    useSelection.checked = loadPref('base.useSelection', '1') !== '0';
    useSelection.addEventListener('change', () => savePref('base.useSelection', useSelection.checked ? '1' : '0'));
  }
  // Old clearspace controls removed - replaced with Akrivi system

  // Base Grid Tiles (single select - radio behavior)
  // Strict radio group: only ONE tile active at a time, all others must revert to OFF state
  const baseTiles = qsa('[data-base]');
  let baseMode = loadPref('base.mode', 'square');
  
  function setBaseMode(mode) {
    if (!mode) return;
    
    baseMode = mode;
    savePref('base.mode', mode);
    
    // CRITICAL: Update ALL tiles to ensure strict radio behavior
    // Remove ON state from all tiles first, then apply to selected one
    baseTiles.forEach(t => {
      const isSelected = t.getAttribute('data-base') === mode;
      
      // Update class and aria state
      if (isSelected) {
        t.classList.add('on');
        t.setAttribute('aria-pressed', 'true');
      } else {
        t.classList.remove('on');
        t.setAttribute('aria-pressed', 'false');
      }
      
      // ALWAYS update icon - ensure proper ON/OFF state (radio behavior)
      // This is critical: tiles must NOT "remember" previous ON states
      const icon = t.querySelector('img.baseGridTileIcon') || t.querySelector('img.baseGridTypeIcon') || t.querySelector('img.tileIcon');
      if (icon) {
        if (isSelected) {
          // Selected tile: Use ON icon asset (blue icon)
          const onSrc = icon.getAttribute('data-icon-on');
          if (onSrc) {
            icon.setAttribute('src', onSrc);
          }
        } else {
          // Unselected tile: ALWAYS use OFF icon asset (white/neutral icon)
          // This ensures tiles don't keep the blue icon after deselection
          const offSrc = icon.getAttribute('data-icon-off');
          if (offSrc) {
            icon.setAttribute('src', offSrc);
          } else {
            // Fallback: if data-icon-off is missing, try to derive from current src
            const currentSrc = icon.getAttribute('src') || '';
            // If current src contains 'ON', replace it to get OFF version
            if (currentSrc.indexOf('ON') !== -1) {
              const derivedOffSrc = currentSrc.replace('ON.svg', '.svg').replace('ON', '');
              icon.setAttribute('src', derivedOffSrc);
            }
          }
        }
      }
    });
  }
  
  // Initialize: set up click handlers
  baseTiles.forEach(t => {
    t.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      const mode = this.getAttribute('data-base');
      if (mode) {
        setBaseMode(mode);
      }
    });
  });
  
  // Initialize state on load
  setBaseMode(baseMode);

  // ============================================
  // Live Preview System
  // ============================================
  
  let livePreviewTimeout = null;
  let livePreviewEnabled = false;
  let objectsGenerated = {
    logo: false,
    clear: false,
    base: false
  };

  function collectLogoGridValues() {
    // CRITICAL: Get grid type directly from active buttons (most reliable source of truth)
    // Supports multi-select as comma-separated string (e.g., "straightLinesGrid,diagonalGrid")
    let gridTypeValue = 'straightLinesGrid'; // default fallback
    const selector = $('lgGridTypeSelector');
    const gridTypeInput = $('lgGridType');
    
    if (selector) {
      // Use fresh query to get current state - ensure DOM is up to date
      const activeTypes = [];
      const activeButtons = selector.querySelectorAll('.shapeBtn.active');
      activeButtons.forEach(btn => {
        const type = btn.getAttribute('data-grid-type');
        if (type && type.trim()) activeTypes.push(type.trim());
      });
      
      // Ensure at least one grid type is selected
      if (activeTypes.length === 0) {
        // No buttons active - activate default
        const defaultBtn = selector.querySelector('[data-grid-type="straightLinesGrid"]');
        if (defaultBtn) {
          defaultBtn.classList.add('active');
          defaultBtn.setAttribute('aria-pressed', 'true');
          activeTypes.push('straightLinesGrid');
        }
      }
      
      // Build comma-separated string from active types
      gridTypeValue = activeTypes.length > 0 ? activeTypes.join(',') : 'straightLinesGrid';
      
      // Always sync to hidden input for consistency and debugging
      if (gridTypeInput) {
        gridTypeInput.value = gridTypeValue;
      }
    } else if (gridTypeInput && gridTypeInput.value && gridTypeInput.value.trim()) {
      // Fallback: use hidden input if selector not found (shouldn't happen in normal operation)
      gridTypeValue = gridTypeInput.value.trim();
    }
    
    return {
      doAnchors: tileOn(logoTiles.anchors) ? 1 : 0,
      doHandles: tileOn(logoTiles.handles) ? 1 : 0,
      doGrid: tileOn(logoTiles.gridlines) ? 1 : 0,
      doOut: tileOn(logoTiles.outlines) ? 1 : 0,
      swAnchors: safeNum($('lgAnchorStroke') && $('lgAnchorStroke').value, 0.5),
      swHandles: safeNum($('lgHandleStroke') && $('lgHandleStroke').value, 0.5),
      swGridlines: safeNum($('lgGridlinesStroke') && $('lgGridlinesStroke').value, 0.5),
      swOutline: safeNum($('lgOutlineStroke') && $('lgOutlineStroke').value, 0.5),
      aSize: safeNum($('lgAnchorSize').value, 3),
      hSize: safeNum($('lgHandleSize').value, 3),
      gridDir: ($('lgGridDir') && $('lgGridDir').value) ? $('lgGridDir').value : 'both',
      aShape: ($('lgAnchorShape') && $('lgAnchorShape').value) ? $('lgAnchorShape').value : 'square',
      hShape: ($('lgHandleShape') && $('lgHandleShape').value) ? $('lgHandleShape').value : 'circle',
      anchorFill: $('lgAnchorFill') && $('lgAnchorFill').checked,
      handleFill: $('lgHandleFill') && $('lgHandleFill').checked,
      gridType: gridTypeValue
    };
  }

  // ============================================
  // Akrivi Clearspace Grid Generator State
  // ============================================
  
  let clearspaceComponents = {
    logoMark: { index: -1, name: null, bounds: null },
    logoType: { index: -1, name: null, bounds: null },
    xSource: { index: -1, name: null, bounds: null }
  };
  
  // Store component names for reliable lookup during generation
  let clearspaceComponentNames = {
    logoMark: null,
    logoType: null,
    xSource: null
  };
  
  /**
   * Collect Akrivi Clearspace Grid values
   */
  function collectClearspaceValues() {
    const lockupRadios = document.querySelectorAll('input[name="csLockupMode"]:checked');
    const lockupMode = lockupRadios.length > 0 ? lockupRadios[0].value : 'horizontal';
    
    const xDefRadios = document.querySelectorAll('input[name="csXDefinition"]:checked');
    const xDefinition = xDefRadios.length > 0 ? xDefRadios[0].value : 'width';
    
    const lockupGapScale = safeNum($('csLockupGapScale').value, 100);
    const clearspaceScale = safeNum($('csClearspaceScale').value, 100);
    
    return {
      logoMarkIndex: clearspaceComponents.logoMark.index,
      logoTypeIndex: clearspaceComponents.logoType.index,
      xSourceIndex: clearspaceComponents.xSource.index,
      logoMarkName: clearspaceComponentNames.logoMark || '',
      logoTypeName: clearspaceComponentNames.logoType || '',
      xSourceName: clearspaceComponentNames.xSource || '',
      lockupMode: lockupMode,
      xDefinition: xDefinition,
      lockupGapScale: lockupGapScale,
      clearspaceScale: clearspaceScale
    };
  }
  
  /**
   * Select component from current selection
   * @param {string} componentType - 'logoMark', 'logoType', or 'xSource'
   */
  function selectComponent(componentType) {
    setStatus('Reading selection...');
    
    ensureJsxLoaded(function() {
      verifyFunction('helpersGetSelectionInfo', function(exists) {
        if (!exists) {
          setStatus('Selection info not available. Try again.');
          return;
        }
        
        const jsx = 'try {' +
          'if ($.global && typeof $.global.helpersGetSelectionInfo === "function") {' +
            '$.global.helpersGetSelectionInfo()' +
          '} else { "Selection info not available." }' +
          '} catch(e) { "ERROR: " + e.toString(); }';
        
        evalJsx(jsx, (res) => {
          const result = String(res);
          
          if (!result || result === 'undefined' || result === 'null') {
            setStatus('Error: No response from Illustrator');
            return;
          }
          
          if (result.indexOf('ERROR') === 0 || result.indexOf('ERR') === 0) {
            setStatus('Error: ' + result);
            return;
          }
          
          try {
            const data = JSON.parse(result);
            
            if (data.error) {
              setStatus('Error: ' + data.error);
              return;
            }
            
            if (data.selectionCount === 0 || !data.items || data.items.length === 0) {
              setStatus('Please select at least one object in Illustrator');
              return;
            }
            
            // Use the first selected object
            const firstItem = data.items[0];
            if (firstItem) {
              const itemName = firstItem.name || 'Selected Object';
              clearspaceComponents[componentType] = {
                index: firstItem.index !== undefined ? firstItem.index : 0,
                name: itemName,
                bounds: {
                  width: firstItem.width || 0,
                  height: firstItem.height || 0
                }
              };
              // Store name for reliable lookup during generation
              clearspaceComponentNames[componentType] = itemName;
              
              updateComponentDisplay(componentType);
              setStatus(componentType === 'logoMark' ? 'Logo Mark' : 
                       componentType === 'logoType' ? 'Logo Type' : 
                       'Exclusion Zone') + ' selected: ' + clearspaceComponents[componentType].name;
            } else {
              setStatus('No valid objects found in selection');
            }
          } catch(e) {
            setStatus('Error parsing selection info: ' + e.message);
            console.error('Selection info parse error:', e, result);
          }
        });
      });
    });
  }
  
  /**
   * Update component display in UI
   */
  function updateComponentDisplay(componentType) {
    const component = clearspaceComponents[componentType];
    const card = $('cs' + componentType.charAt(0).toUpperCase() + componentType.slice(1) + 'Card');
    const status = $('cs' + componentType.charAt(0).toUpperCase() + componentType.slice(1) + 'Status');
    const preview = $('cs' + componentType.charAt(0).toUpperCase() + componentType.slice(1) + 'Preview');
    
    if (component.index >= 0 && component.name) {
      if (status) status.textContent = component.name;
      if (card) card.classList.add('active');
      if (preview) {
        preview.innerHTML = '<span class="component-preview-text">' + 
          component.name + '<br>' +
          component.bounds.width.toFixed(1) + ' × ' + component.bounds.height.toFixed(1) + ' px' +
          '</span>';
      }
    } else {
      if (status) status.textContent = 'Not selected';
      if (card) card.classList.remove('active');
      if (preview) {
        const defaultText = componentType === 'logoMark' ? 'Icon, Symbol, or Logomark' :
                           componentType === 'logoType' ? 'Wordmark or Typography' :
                           'Defines X-unit';
        preview.innerHTML = '<span class="component-preview-text">' + defaultText + '</span>';
      }
    }
  }
  
  /**
   * Update all component displays
   */
  function updateAllComponentDisplays() {
    updateComponentDisplay('logoMark');
    updateComponentDisplay('logoType');
    updateComponentDisplay('xSource');
  }
  
  /**
   * Reset Clearspace Grid Generator
   */
  function resetClearspace() {
    clearspaceComponents = {
      logoMark: { index: -1, name: null, bounds: null },
      logoType: { index: -1, name: null, bounds: null },
      xSource: { index: -1, name: null, bounds: null }
    };
    
    clearspaceComponentNames = {
      logoMark: null,
      logoType: null,
      xSource: null
    };
    
    updateAllComponentDisplays();
    
    // Reset form values
    if ($('csLockupGapScale')) $('csLockupGapScale').value = '100';
    if ($('csClearspaceScale')) $('csClearspaceScale').value = '100';
    
    const horizontalRadio = $('csLockupHorizontal');
    if (horizontalRadio) horizontalRadio.checked = true;
    
    const widthRadio = $('csXDefinitionWidth');
    if (widthRadio) widthRadio.checked = true;
    
    // Reset scaling buttons
    document.querySelectorAll('.scaling-btn').forEach(btn => {
      btn.classList.remove('active');
    });
    
    setStatus('Clearspace Grid Generator reset');
  }

  function collectBaseGridValues() {
    return {
      mode: baseMode,
      size: safeNum($('gridSize').value, 40),
      stroke: safeNum($('strokeW') && $('strokeW').value, 0.5),
      useSel: ($('useSelection') && $('useSelection').checked) ? 1 : 0
    };
  }

  /**
   * CRUCIAL: When Global Color or Global Stroke changes, ONLY update existing elements in place.
   * DO NOT call generate() or clean(). This function is the ONLY path used by Global Color / Global Stroke listeners.
   */
  function applyGlobalStyleToExistingElementsOnly() {
    if (currentView === 'logo' && objectsGenerated.logo) {
      const vals = collectLogoGridValues();
      if (vals.doAnchors || vals.doHandles || vals.doGrid || vals.doOut) {
        updateLogoGridLive(vals);
      }
      return;
    }
    if (currentView === 'base' && objectsGenerated.base) {
      const vals = collectBaseGridValues();
      updateBaseGridLive(vals);
    }
  }

  /**
   * Live preview: only updates existing items (updateVisuals). Never calls clean() or generate().
   */
  function triggerLivePreview() {
    if (!livePreviewEnabled) return;
    if (currentView === 'logo' && !objectsGenerated.logo) return;
    if (currentView === 'clear' && !objectsGenerated.clear) return;
    if (currentView === 'base' && !objectsGenerated.base) return;
    if (livePreviewTimeout) clearTimeout(livePreviewTimeout);
    livePreviewTimeout = setTimeout(() => {
      if (currentView === 'logo') {
        const vals = collectLogoGridValues();
        if (vals.doAnchors || vals.doHandles || vals.doGrid || vals.doOut) {
          updateLogoGridLive(vals);
        }
      } else if (currentView === 'clear') {
        // Clear Space does NOT use live preview
      } else if (currentView === 'base') {
        const vals = collectBaseGridValues();
        updateBaseGridLive(vals);
      }
    }, 300);
  }

  /**
   * Run live update for ONE Logo Grid feature only. No side effects on other features.
   * feature: 'anchors' | 'handles' | 'outline' | 'gridlines'
   */
  function runLogoGridLiveUpdate(feature) {
    if (!livePreviewEnabled || currentView !== 'logo' || !objectsGenerated.logo) return;
    const vals = collectLogoGridValues();
    const color = getGlobalColor();
    const sw = getGlobalStrokeWeight();
    const swOutline = Number(vals.swOutline) || sw;
    const run = function (fnName, args) {
      const jsx = 'try { if ($.global && typeof $.global.' + fnName + ' === "function") { $.global.' + fnName + '(' + args + '); } } catch(e) {}';
      evalJsx(jsx, () => {});
    };
    if (feature === 'anchors') {
      run('helpersUpdateLogoGridAnchors', [
        String(color.r), String(color.g), String(color.b),
        String(vals.aSize), '"' + (vals.aShape || 'square') + '"',
        vals.anchorFill ? 'true' : 'false',
        String(vals.swAnchors)
      ].join(','));
    } else if (feature === 'handles') {
      run('helpersUpdateLogoGridHandles', [
        String(color.r), String(color.g), String(color.b),
        String(vals.hSize), '"' + (vals.hShape || 'circle') + '"',
        String(vals.swHandles),
        vals.handleFill ? 'true' : 'false'
      ].join(','));
    } else if (feature === 'outline' && vals.doOut) {
      run('helpersUpdateLogoGridOutline', [
        String(color.r), String(color.g), String(color.b),
        String(swOutline)
      ].join(','));
    } else if (feature === 'gridlines') {
      run('helpersUpdateLogoGridGridlines', [
        String(color.r), String(color.g), String(color.b),
        String(vals.swGridlines)
      ].join(','));
    }
  }

  /**
   * updateVisuals only: find existing objects by name (Anchor, HandleLine, HandleDot, Outline, Gridline) and update in place.
   * Never calls clean() or generate(). No deletion, no creation. Used by Apply Global Style and full refresh.
   */
  function updateLogoGridLive(vals) {
    const color = getGlobalColor();
    const sw = getGlobalStrokeWeight();
    const swOutline = Number(vals.swOutline) || sw;
    const run = function (fnName, args) {
      const jsx = 'try { if ($.global && typeof $.global.' + fnName + ' === "function") { $.global.' + fnName + '(' + args + '); } } catch(e) {}';
      evalJsx(jsx, () => {});
    };
    if (vals.doAnchors) {
      run('helpersUpdateLogoGridAnchors', [
        String(color.r), String(color.g), String(color.b),
        String(vals.aSize), '"' + (vals.aShape || 'square') + '"',
        vals.anchorFill ? 'true' : 'false',
        String(vals.swAnchors)
      ].join(','));
    }
    if (vals.doHandles) {
      run('helpersUpdateLogoGridHandles', [
        String(color.r), String(color.g), String(color.b),
        String(vals.hSize), '"' + (vals.hShape || 'circle') + '"',
        String(vals.swHandles),
        vals.handleFill ? 'true' : 'false'
      ].join(','));
    }
    if (vals.doOut) {
      run('helpersUpdateLogoGridOutline', [
        String(color.r), String(color.g), String(color.b),
        String(swOutline)
      ].join(','));
    }
    if (vals.doGrid) {
      run('helpersUpdateLogoGridGridlines', [
        String(color.r), String(color.g), String(color.b),
        String(vals.swGridlines)
      ].join(','));
    }
  }

  // Clearspace live update: only update stroke/color on existing Clearspace_Grid_Output. No delete, no re-create.
  // Clearspace always uses fixed gray; Global Color in Settings does not affect it.
  function updateClearspaceLive() {
    if (!objectsGenerated.clear) return;
    verifyFunction('helpersUpdateClearspaceVisuals', function(exists) {
      if (!exists) return;
      const strokeW = safeNum($('csStroke') && $('csStroke').value, 0.5);
      const colorClearspace = CLEARSPACE_GRAY;
      const jsx = 'try { if ($.global && typeof $.global.helpersUpdateClearspaceVisuals === "function") { $.global.helpersUpdateClearspaceVisuals(' +
        String(strokeW) + ', ' + colorClearspace.r + ', ' + colorClearspace.g + ', ' + colorClearspace.b + '); } } catch(e) {}';
      evalJsx(jsx, () => {});
    });
  }

  /**
   * Update Base Grid visuals only (stroke/color on existing Group_BaseGrid items). No deletion, no regeneration.
   */
  function updateBaseGridLive(vals) {
    const run = function(fnName, args) {
      const jsx = 'try { if ($.global && typeof $.global.' + fnName + ' === "function") { $.global.' + fnName + '(' + args + '); } } catch(e) {}';
      evalJsx(jsx, () => {});
    };
    run('helpersUpdateBaseGridVisuals', [String(vals.stroke)].join(','));
  }

  // Setup live preview listeners for all inputs (NOT tiles - tiles only toggle state)
  // Logo Grid: each input only runs its feature updater (anchors/handles/outline/gridlines) - no cross-feature side effects
  function setupLivePreviewListeners() {
    const logoInputToFeature = {
      lgAnchorSize: 'anchors',
      lgAnchorStroke: 'anchors',
      lgHandleSize: 'handles',
      lgHandleStroke: 'handles',
      lgOutlineStroke: 'outline',
      lgGridlinesStroke: 'gridlines',
      lgGridType: 'gridlines'
    };
    Object.keys(logoInputToFeature).forEach(id => {
      const el = $(id);
      const feature = logoInputToFeature[id];
      if (el && feature) {
        el.addEventListener('input', function () { runLogoGridLiveUpdate(feature); });
        el.addEventListener('change', function () { runLogoGridLiveUpdate(feature); });
      }
    });
    const anchorFillEl = $('lgAnchorFill');
    if (anchorFillEl) {
      anchorFillEl.addEventListener('change', function () { runLogoGridLiveUpdate('anchors'); });
    }
    const handleFillEl = $('lgHandleFill');
    if (handleFillEl) {
      handleFillEl.addEventListener('change', function () { runLogoGridLiveUpdate('handles'); });
    }

    // NOTE: Logo Grid tiles (Anchors, Handles, etc.) do NOT trigger live preview
    // They only toggle active state - generation happens on "Generate" button click

    // Akrivi Clearspace doesn't use live preview
    // All inputs are handled by their own event listeners in initializeClearspaceUI()

    // Base Grid inputs
    const baseInputs = ['gridSize', 'strokeW', 'useSelection'];
    baseInputs.forEach(id => {
      const el = $(id);
      if (el) {
        el.addEventListener('input', triggerLivePreview);
        el.addEventListener('change', triggerLivePreview);
      }
    });

    // NOTE: Base Grid tiles do NOT trigger live preview
    // They only toggle active state - generation happens on "Generate" button click
  }

  // ============================================
  // Action Handlers
  // ============================================
  
  /**
   * Setup action button handlers (Generate, Clean)
   */
  function setupActionButtons() {
    const btnCleanEl = $('btnClean');
    if (btnCleanEl) {
      btnCleanEl.addEventListener('click', () => {
        setStatus('Cleaning...');
        // ONLY the Clean button calls helpersCleanupByMode. Color/stroke changes never call clean().
        var mode = currentView === 'clear' ? 'clearspace' : currentView; // 'logo' | 'clearspace' | 'base'
        var script = 'try { ' +
          'if ($.global && typeof $.global.helpersCleanupByMode === "function") { ' +
            '$.global.helpersCleanupByMode("' + mode + '")' +
          '} else { "Clean not available." }' +
          '} catch(e) { "ERROR: " + e.toString(); }';
        evalJsx(script, (res) => setStatus(String(res)));
      });
    }

    /**
     * Generate button: ONLY place that calls helpersGenerateLogoGrid / helpersGenerateBaseGrid (creates new items).
     * Color/stroke changes never call generate(); they only call updateLogoGridLive / updateBaseGridLive (updateVisuals).
     */
    const btnPrimaryEl = $('btnPrimary');
    if (!btnPrimaryEl) {
      console.error('btnPrimary element not found');
      return;
    }
    
    btnPrimaryEl.addEventListener('click', () => {
    if (currentView === 'logo') {
      // multi-select modules
      const doAnchors = tileOn(logoTiles.anchors) ? 1 : 0;
      const doHandles = tileOn(logoTiles.handles) ? 1 : 0;
      const doGrid    = tileOn(logoTiles.gridlines) ? 1 : 0;
      const doOut     = tileOn(logoTiles.outlines) ? 1 : 0;
      const swAnchors  = 0.5;
      const swHandles  = getGlobalStrokeWeight();
      const swGridlines = getGlobalStrokeWeight();
      const swOutline  = getGlobalStrokeWeight();
      const aSize = safeNum($('lgAnchorSize').value, 3);
      const hSize = safeNum($('lgHandleSize').value, 3); // INDEPENDENT handle size
      // CRITICAL: Sync grid type from active buttons to hidden input before collecting values
      // This ensures the hidden input always reflects the current button state
      // This is the source of truth for grid type selection
      const gridTypeSelector = $('lgGridTypeSelector');
      const gridTypeInput = $('lgGridType');
      if (gridTypeSelector && gridTypeInput) {
        const activeTypes = [];
        const buttons = gridTypeSelector.querySelectorAll('.shapeBtn.active');
        buttons.forEach(btn => {
          const type = btn.getAttribute('data-grid-type');
          if (type) activeTypes.push(type);
        });
        
        // Ensure at least one grid type is always selected
        if (activeTypes.length === 0) {
          // No buttons active - activate default
          const defaultBtn = gridTypeSelector.querySelector('[data-grid-type="straightLinesGrid"]');
          if (defaultBtn) {
            defaultBtn.classList.add('active');
            defaultBtn.setAttribute('aria-pressed', 'true');
            activeTypes.push('straightLinesGrid');
          }
        }
        
        // Always set the hidden input to the active types
        const gridTypesString = activeTypes.length > 0 ? activeTypes.join(',') : 'straightLinesGrid';
        gridTypeInput.value = gridTypesString;
      }

      setStatus('Generating Logo Grid...');
      
      // CRITICAL: Collect all values including gridType from active buttons
      const vals = collectLogoGridValues();
      
      // CRITICAL VERIFICATION: Ensure gridType is set and valid
      // This is the final check before passing to JSX - read directly from buttons if needed
      if (!vals.gridType || typeof vals.gridType !== 'string' || vals.gridType.trim() === '') {
        // Fallback: read directly from buttons one more time as absolute source of truth
        const gridTypeSelector = $('lgGridTypeSelector');
        if (gridTypeSelector) {
          const activeTypes = [];
          const buttons = gridTypeSelector.querySelectorAll('.shapeBtn.active');
          buttons.forEach(btn => {
            const type = btn.getAttribute('data-grid-type');
            if (type && type.trim()) activeTypes.push(type.trim());
          });
          if (activeTypes.length > 0) {
            vals.gridType = activeTypes.join(',');
          } else {
            // No buttons active - activate default and use it
            const defaultBtn = gridTypeSelector.querySelector('[data-grid-type="straightLinesGrid"]');
            if (defaultBtn) {
              defaultBtn.classList.add('active');
              defaultBtn.setAttribute('aria-pressed', 'true');
            }
            vals.gridType = 'straightLinesGrid';
          }
        } else {
          vals.gridType = 'straightLinesGrid';
        }
      }
      
      // Final validation: ensure gridType is a non-empty string
      if (!vals.gridType || typeof vals.gridType !== 'string') {
        vals.gridType = 'straightLinesGrid';
      }
      
      verifyFunction('helpersGenerateLogoGrid', function(exists) {
        if (!exists) {
          setStatus('Logo Grid not available.');
          return;
        }
        
        const color = getGlobalColor();
        const gridTypeParam = (vals.gridType && vals.gridType.trim()) ? vals.gridType.trim() : 'straightLinesGrid';
        
        const jsx = 'try {' +
          'if ($.global && typeof $.global.helpersGenerateLogoGrid === "function") {' +
            '$.global.helpersGenerateLogoGrid(' +
              (vals.doAnchors ? 'true' : 'false') + ',' +
              (vals.doHandles ? 'true' : 'false') + ',' +
              (vals.doOut ? 'true' : 'false') + ',' +
              (vals.doGrid ? 'true' : 'false') + ',' +
              String(vals.swAnchors) + ',' +
              String(vals.swHandles) + ',' +
              String(vals.swOutline) + ',' +
              String(vals.swGridlines) + ',' +
              String(vals.aSize) + ',' +
              String(vals.hSize) + ',' +
            '"' + vals.gridDir + '"' + ',' +
            '"' + vals.aShape + '"' + ',' +
            '"' + vals.hShape + '"' + ',' +
            (vals.anchorFill ? 'true' : 'false') + ',' +
            (vals.handleFill ? 'true' : 'false') + ',' +
            '"' + gridTypeParam + '"' + ',' +
            'false' + ',' +
              String(color.r) + ',' + String(color.g) + ',' + String(color.b) + ',' +
              String(color.r) + ',' + String(color.g) + ',' + String(color.b) + ',' +
              String(color.r) + ',' + String(color.g) + ',' + String(color.b) +
            ');' +
          '} else { "Logo Grid not available." }' +
          '} catch(e) { "ERROR: " + e.toString(); }';
        
        evalJsx(jsx, (res) => {
          const result = String(res);
          // Mark as generated and enable live preview for this view
          if (result.indexOf('ERROR') !== 0 && result.indexOf('ERR') !== 0 && result.indexOf('not available') === -1) {
            objectsGenerated.logo = true;
            livePreviewEnabled = true;
            updateLivePreviewIndicator();
          }
          setStatus((result.indexOf('ERROR') === 0 || result.indexOf('ERR') === 0 || result.indexOf('not available') !== -1) ? result : 'Logo Grid generated');
        });
      });
      return;
    }

    if (currentView === 'clear') {
      // Use the new minimal Clearspace generation logic
      runClearspaceGeneratorFromMain();
      return;
    }

    if (currentView === 'base') {
      const vals = collectBaseGridValues();
      
      setStatus('Generating Base Grid...');
      verifyFunction('helpersGenerateBaseGrid', function(exists) {
        if (!exists) {
          setStatus('Base Grid not available.');
          return;
        }
        const jsx = 'try {' +
          'if ($.global && typeof $.global.helpersGenerateBaseGrid === "function") {' +
            '$.global.helpersGenerateBaseGrid(' +
              '"' + vals.mode + '"' + ',' +
              vals.size + ',' +
              vals.stroke + ',' +
              vals.useSel + ',' +
              'false' + // updateMode = false for initial generation
            ')' +
          '} else { "Base Grid not available." }' +
          '} catch(e) { "ERROR: " + e.toString(); }';
        evalJsx(jsx, (res) => {
          setStatus(String(res));
          // Mark as generated and enable live preview for this view
          if (String(res).indexOf('ERROR') !== 0 && String(res).indexOf('ERR') !== 0) {
            objectsGenerated.base = true;
            livePreviewEnabled = true;
            updateLivePreviewIndicator();
          }
        });
      });
    }
    });
  }

  // ============================================
  // Application Initialization
  // ============================================
  
  /**
   * Initialize application on load
   */
  function initializeApp() {
    refreshBoundControls();
    showView(currentView);
    setStatus('Ready');
    
    // Setup action buttons (Generate, Clean)
    setupActionButtons();
    
    // Setup live preview listeners (but don't enable until first Generate click)
    setTimeout(() => {
      setupLivePreviewListeners();
      // livePreviewEnabled remains false until first successful generation
    }, 500); // Small delay to ensure all event handlers are attached
  }
  
  // ============================================
  // Clearspace Generator UI Initialization
  // ============================================
  
  /**
   * Initialize Clearspace Grid Generator UI
   */
  // Minimal Clearspace UI state
  let clearspaceXRefIndex = -1;
  let clearspaceXRefInfo = null;
  let clearSpacePercent = 100; // Default: 100%

  function initializeClearspaceUI() {
    // Pick X Source button
    const pickXSourceBtn = $('csPickXSource');
    if (pickXSourceBtn) {
      pickXSourceBtn.addEventListener('click', handlePickXSource);
    }

    // Stroke (px) — pref + live update for existing grid
    const csStrokeEl = $('csStroke');
    if (csStrokeEl) {
      csStrokeEl.value = loadPref('cs.stroke', '0.5');
      csStrokeEl.addEventListener('input', function () {
        savePref('cs.stroke', csStrokeEl.value);
        updateClearspaceLive();
      });
      csStrokeEl.addEventListener('change', function () {
        savePref('cs.stroke', csStrokeEl.value);
        updateClearspaceLive();
      });
    }
    
    // Percentage selection buttons - NO live preview (only updates on Generate)
    const percentButtons = ['csPercent100', 'csPercent50', 'csPercent25'];
    percentButtons.forEach(btnId => {
      const btn = $(btnId);
      if (btn) {
        btn.addEventListener('click', function() {
          // Remove active class from all buttons
          percentButtons.forEach(id => {
            const b = $(id);
            if (b) b.classList.remove('active');
          });
          
          // Add active class to clicked button
          this.classList.add('active');
          
          // Store selected percentage
          clearSpacePercent = parseInt(this.getAttribute('data-percent'), 10);
          
          // NO live preview - Clear Space only updates on Generate button click
        });
      }
    });
  }

  function handlePickXSource() {
    const errorMsg = $('csErrorMsg');
    const readout = $('csReadout');
    const xSourceType = $('csXSourceType');
    const xSourceSize = $('csXSourceSize');
    const xSourceName = $('csXSourceName');
    
    // Always overwrite previous selection: forget old X-Source data so the new pick is the only source
    clearspaceXRefIndex = -1;
    clearspaceXRefInfo = null;
    
    // Hide previous errors/readouts
    if (errorMsg) errorMsg.classList.add('hidden');
    if (readout) readout.classList.add('hidden');
    
    setStatus('Reading selection...');
    
    // Command Illustrator to return current active selection only (no cached data)
    const jsx = 'try { ' +
      'if ($.global && typeof $.global.helpersGetSelectionInfoMinimal === "function") { ' +
        '$.global.helpersGetSelectionInfoMinimal()' +
      '} else { "Selection not available." }' +
      '} catch(e) { "ERROR: " + e.toString(); }';
    
    evalJsx(jsx, (res) => {
      try {
        const result = JSON.parse(String(res));
        
        if (!result.hasSelection || !result.items || result.items.length === 0) {
          if (errorMsg) {
            errorMsg.textContent = 'Select an object in Illustrator first';
            errorMsg.classList.remove('hidden');
          }
          setStatus('No selection found');
          clearspaceXRefIndex = -1;
          clearspaceXRefInfo = null;
          return;
        }
        
        // Use recommended index (last selected) or first item
        const index = result.recommendedIndex >= 0 ? result.recommendedIndex : 0;
        const item = result.items[index];
        
        if (!item) {
          if (errorMsg) {
            errorMsg.textContent = 'Could not read selection item';
            errorMsg.classList.remove('hidden');
          }
          setStatus('Error reading selection');
          return;
        }
        
        // Store index
        clearspaceXRefIndex = index;
        clearspaceXRefInfo = item;
        
        // Set as X-REF if auto-rename is enabled
        const autoRename = $('csAutoRename');
        if (autoRename && autoRename.checked) {
          setXRefName(index, item);
        } else {
          updateXSourceReadout(item);
          setStatus('Source Updated: ' + (item.name || item.typename || 'X Source'));
        }
      } catch (e) {
        if (errorMsg) {
          errorMsg.textContent = 'Error: ' + e.toString();
          errorMsg.classList.remove('hidden');
        }
        setStatus('Error: ' + e.toString());
      }
    });
  }

  function setXRefName(index, itemInfo) {
    const jsx = 'try { ' +
      'if ($.global && typeof $.global.helpersSetXRef === "function") { ' +
        '$.global.helpersSetXRef(' + index + ')' +
      '} else { "Source update not available." }' +
      '} catch(e) { "ERROR: " + e.toString(); }';
    
    evalJsx(jsx, (res) => {
      try {
        const result = JSON.parse(String(res));
        
        if (result.success) {
          updateXSourceReadout({
            typename: result.typename,
            name: result.name,
            widthPt: result.widthPt,
            heightPt: result.heightPt
          });
          setStatus('Source Updated: ' + (result.name || 'X-REF'));
        } else {
          const errorMsg = $('csErrorMsg');
          if (errorMsg) {
            errorMsg.textContent = result.error || 'Failed to set X-REF';
            errorMsg.classList.remove('hidden');
          }
          setStatus('Error setting X-REF');
        }
      } catch (e) {
        const errorMsg = $('csErrorMsg');
        if (errorMsg) {
          errorMsg.textContent = 'Error: ' + e.toString();
          errorMsg.classList.remove('hidden');
        }
      }
    });
  }

  function updateXSourceReadout(item) {
    const readout = $('csReadout');
    const xSourceType = $('csXSourceType');
    const xSourceSize = $('csXSourceSize');
    const xSourceName = $('csXSourceName');
    
    if (xSourceType) xSourceType.textContent = item.typename || '—';
    if (xSourceSize) xSourceSize.textContent = (item.widthPt || 0).toFixed(1) + ' × ' + (item.heightPt || 0).toFixed(1) + ' px';
    if (xSourceName) xSourceName.textContent = item.name || '(unnamed)';
    
    if (readout) readout.classList.remove('hidden');
  }

  function handleGenerateClearspace() {
    const statusMsg = $('csStatusMsg');
    const errorMsg = $('csErrorMsg');
    
    // Hide previous messages
    if (statusMsg) statusMsg.classList.add('hidden');
    if (errorMsg) errorMsg.classList.add('hidden');
    
    // Check if X-REF is set
    if (clearspaceXRefIndex < 0) {
      // Try to use current selection if auto-rename is on
      const autoRename = $('csAutoRename');
      if (autoRename && autoRename.checked) {
        // Get selection and use first item
        const jsx = 'try { ' +
          'if ($.global && typeof $.global.helpersGetSelectionInfoMinimal === "function") { ' +
            '$.global.helpersGetSelectionInfoMinimal()' +
          '} else { "Selection not available." }' +
          '} catch(e) { "ERROR: " + e.toString(); }';
        
        evalJsx(jsx, (res) => {
          try {
            const result = JSON.parse(String(res));
            if (result.hasSelection && result.items && result.items.length > 0) {
              const index = result.recommendedIndex >= 0 ? result.recommendedIndex : 0;
              setXRefName(index, result.items[index]);
              // Wait a moment then generate
              setTimeout(() => runClearspaceGenerator(), 300);
            } else {
              if (errorMsg) {
                errorMsg.textContent = 'Please select an object and click "Pick X Source" first';
                errorMsg.classList.remove('hidden');
              }
              setStatus('No selection found');
            }
          } catch (e) {
            if (errorMsg) {
              errorMsg.textContent = 'Error: ' + e.toString();
              errorMsg.classList.remove('hidden');
            }
          }
        });
        return;
      } else {
        if (errorMsg) {
          errorMsg.textContent = 'Please select an object and click "Pick X Source" first';
          errorMsg.classList.remove('hidden');
        }
        setStatus('X Source not selected');
        return;
      }
    }
    
    runClearspaceGenerator();
  }

  function runClearspaceGenerator() {
    runClearspaceGeneratorFromMain();
  }

  function runClearspaceGeneratorFromMain() {
    const statusMsg = $('csStatusMsg');
    const errorMsg = $('csErrorMsg');
    
    // Hide previous messages
    if (statusMsg) statusMsg.classList.add('hidden');
    if (errorMsg) errorMsg.classList.add('hidden');
    
    // Check if X-REF is set
    if (clearspaceXRefIndex < 0) {
      // Try to use current selection if auto-rename is on
      const autoRename = $('csAutoRename');
      if (autoRename && autoRename.checked) {
        // Get selection and use first item
        const jsx = 'try { ' +
          'if ($.global && typeof $.global.helpersGetSelectionInfoMinimal === "function") { ' +
            '$.global.helpersGetSelectionInfoMinimal()' +
          '} else { "Selection not available." }' +
          '} catch(e) { "ERROR: " + e.toString(); }';
        
        evalJsx(jsx, (res) => {
          try {
            const result = JSON.parse(String(res));
            if (result.hasSelection && result.items && result.items.length > 0) {
              const index = result.recommendedIndex >= 0 ? result.recommendedIndex : 0;
              setXRefName(index, result.items[index]);
              // Wait a moment then generate
              setTimeout(() => runClearspaceGeneratorWithParams(), 300);
            } else {
              if (errorMsg) {
                errorMsg.textContent = 'Please select an object and click "Pick X Source" first';
                errorMsg.classList.remove('hidden');
              }
              setStatus('No selection found');
            }
          } catch (e) {
            if (errorMsg) {
              errorMsg.textContent = 'Error: ' + e.toString();
              errorMsg.classList.remove('hidden');
            }
            setStatus('Error: ' + e.toString());
          }
        });
        return;
      } else {
        if (errorMsg) {
          errorMsg.textContent = 'Please select an object and click "Pick X Source" first';
          errorMsg.classList.remove('hidden');
        }
        setStatus('X Source not selected');
        return;
      }
    }
    
    runClearspaceGeneratorWithParams();
  }

  function runClearspaceGeneratorWithParams() {
    const statusMsg = $('csStatusMsg');
    const errorMsg = $('csErrorMsg');
    
    setStatus('Generating Clearspace Grid...');
    
    // Parameters (Stroke from csStroke input)
    const xMode = '3'; // AUTO (square-safe)
    const strokeW = safeNum($('csStroke') && $('csStroke').value, 0.5);
    const cornerOpacity = '18';
    const clearspacePercent = clearSpacePercent || 100; // Use stored percentage
    const colorClearspace = CLEARSPACE_GRAY; // Clearspace always gray; Global Color does not apply
      const jsx = 'try { ' +
        'if ($.global && typeof $.global.helpersRunClearspace === "function") { ' +
          '$.global.helpersRunClearspace("' + xMode + '", ' + String(strokeW) + ', ' + cornerOpacity + ', ' + clearspacePercent + ', false, ' +
          colorClearspace.r + ', ' + colorClearspace.g + ', ' + colorClearspace.b + ')' +
        '} else { "Clearspace not available." }' +
        '} catch(e) { "ERROR: " + e.toString(); }';
    
    evalJsx(jsx, (res) => {
      const result = String(res);
      
      if (result.indexOf('SUCCESS') === 0) {
        if (statusMsg) {
          statusMsg.textContent = 'Done';
          statusMsg.classList.remove('hidden');
        }
        setStatus('Clearspace grid generated');
        objectsGenerated.clear = true;
        livePreviewEnabled = true;
        updateLivePreviewIndicator();
        
        // Hide after 3 seconds
        setTimeout(() => {
          if (statusMsg) statusMsg.classList.add('hidden');
        }, 3000);
      } else {
        if (errorMsg) {
          errorMsg.textContent = result.replace(/^(ERROR|ERR):\s*/, '');
          errorMsg.classList.remove('hidden');
        }
        setStatus('Error: ' + result);
      }
    });
  }
  
  /**
   * Update scaling button active state
   */
  function updateScalingButtonState(inputId, value) {
    const targetButtons = document.querySelectorAll(`.scaling-btn[data-target="${inputId}"]`);
    targetButtons.forEach(btn => {
      const btnScale = parseFloat(btn.getAttribute('data-scale'));
      btn.classList.toggle('active', Math.abs(btnScale - value) < 1);
    });
  }
  
  // Initialize when DOM is ready
  initializeApp();
  initializeClearspaceUI();
  
  // Initial Live Preview indicator update
  setTimeout(() => {
    updateLivePreviewIndicator();
  }, 100);

  // Update check (non-blocking, runs after panel is ready)
  setTimeout(checkForUpdates, 2500);
})();
