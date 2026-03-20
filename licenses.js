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

