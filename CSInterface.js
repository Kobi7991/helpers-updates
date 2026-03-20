/* Minimal CSInterface.js for Adobe CEP panels
 * Provides:
 *  - window.CSInterface
 *  - window.SystemPath
 * Works in CEP. Degrades gracefully in a normal browser.
 */

(function (global) {
  'use strict';

  var cep = global.__adobe_cep__;

  // Matches Adobe CEP numeric enum
  var SystemPath = {
    USER_DATA: 0,
    COMMON_FILES: 1,
    MY_DOCUMENTS: 2,
    APPLICATION: 3,
    EXTENSION: 4,
    HOST_APPLICATION: 5
  };

  function CSInterface() {}

  CSInterface.prototype.getSystemPath = function (pathType) {
    if (!cep || !cep.getSystemPath) return "";
    return cep.getSystemPath(pathType);
  };

  CSInterface.prototype.evalScript = function (script, callback) {
    if (!cep || !cep.evalScript) {
      if (typeof callback === "function") callback("EvalScript unavailable (not running in CEP)." );
      return;
    }
    cep.evalScript(script, callback);
  };

  CSInterface.prototype.openURLInDefaultBrowser = function (url) {
    if (cep && cep.openURLInDefaultBrowser) return cep.openURLInDefaultBrowser(url);
    try { global.open(url, "_blank"); } catch (e) {}
  };

  CSInterface.prototype.addEventListener = function (type, listener) {
    if (cep && cep.addEventListener) return cep.addEventListener(type, listener);
  };

  CSInterface.prototype.removeEventListener = function (type, listener) {
    if (cep && cep.removeEventListener) return cep.removeEventListener(type, listener);
  };

  CSInterface.prototype.dispatchEvent = function (event) {
    if (cep && cep.dispatchEvent) return cep.dispatchEvent(event);
  };

  global.SystemPath = global.SystemPath || SystemPath;
  global.CSInterface = global.CSInterface || CSInterface;

})(this);
