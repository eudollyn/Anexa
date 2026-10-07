/* Anexa - código compartilhado entre o painel e os eventos automáticos. */
(function (g) {
  "use strict";

  var DEFAULTS = {
    enabled: true,        // automação ligada
    autoAttach: true,     // anexar ao abrir e-mail novo
    autoSend: true,       // enviar quando houver destinatário
    sendReplies: false,   // também enviar respostas/encaminhamentos sozinho
    requireValid: true,   // todos os destinatários precisam ter e-mail válido
    validateFile: true,   // só envia se o anexo ainda estiver no e-mail
    delayMs: 3000,        // tempo sem mudanças nos destinatários antes de enviar
    fileName: "",
    fileSize: 0
  };
  var SETTINGS_KEY = "anexa";
  var LOG_KEY = "anexa.log";

  // Promessa a partir das APIs com callback do Office.js.
  function op(target, method) {
    var args = Array.prototype.slice.call(arguments, 2);
    return new Promise(function (resolve, reject) {
      args.push(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
        else reject(r.error || new Error(method + " falhou"));
      });
      target[method].apply(target, args);
    });
  }

  function loadSettings() {
    var saved = null;
    try { saved = Office.context.roamingSettings.get(SETTINGS_KEY); } catch (e) { }
    var s = {};
    for (var k in DEFAULTS) s[k] = DEFAULTS[k];
    if (saved) for (var k2 in saved) if (k2 in DEFAULTS) s[k2] = saved[k2];
    s.delayMs = Math.min(15000, Math.max(1000, Number(s.delayMs) || DEFAULTS.delayMs));
    return s;
  }

  function saveSettings(s) {
    Office.context.roamingSettings.set(SETTINGS_KEY, s);
    return op(Office.context.roamingSettings, "saveAsync");
  }

  // O arquivo fica no IndexedDB (as configurações de roaming só aceitam 32 KB).
  function openDb() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open("anexa", 1);
      req.onupgradeneeded = function () { req.result.createObjectStore("kv"); };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }
  function idb(mode, fn) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction("kv", mode);
        var req = fn(tx.objectStore("kv"));
        tx.oncomplete = function () { db.close(); resolve(req && req.result); };
        tx.onerror = function () { db.close(); reject(tx.error); };
      });
    });
  }
  function getFile() { return idb("readonly", function (st) { return st.get("file"); }); }
  function setFile(f) { return idb("readwrite", function (st) { return st.put(f, "file"); }); }
  function clearFile() { return idb("readwrite", function (st) { return st.delete("file"); }); }

  function log(message, ok) {
    try {
      var list = JSON.parse(localStorage.getItem(LOG_KEY) || "[]");
      list.unshift({ t: Date.now(), m: message, ok: ok !== false });
      localStorage.setItem(LOG_KEY, JSON.stringify(list.slice(0, 60)));
    } catch (e) { }
    try { console.log("[Anexa] " + message); } catch (e2) { }
  }
  function getLog() {
    try { return JSON.parse(localStorage.getItem(LOG_KEY) || "[]"); } catch (e) { return []; }
  }
  function clearLog() { try { localStorage.removeItem(LOG_KEY); } catch (e) { } }

  var EMAIL = /^[^\s@<>()",;]+@[^\s@<>()",;]+\.[^\s@<>()",;]{2,}$/;
  function isValidEmail(a) { return EMAIL.test(String(a || "").trim()); }

  function canSend() {
    try {
      return Office.context.requirements.isSetSupported("Mailbox", "1.15") &&
        typeof Office.context.mailbox.item.sendAsync === "function";
    } catch (e) { return false; }
  }

  function hasAttachment(item, name) {
    return op(item, "getAttachmentsAsync").then(function (atts) {
      var n = String(name).toLowerCase();
      return (atts || []).some(function (a) { return String(a.name).toLowerCase() === n; });
    });
  }

  // Adiciona o arquivo padrão ao item em composição, sem duplicar.
  function attachDefault(item) {
    return getFile().then(function (f) {
      if (!f || !f.base64) throw new Error("nenhum arquivo padrão salvo. Abra o painel do Anexa e escolha o arquivo");
      return hasAttachment(item, f.name).then(function (already) {
        if (already) return { name: f.name, added: false };
        return op(item, "addFileAttachmentFromBase64Async", f.base64, f.name, { isInline: false })
          .then(function () { return { name: f.name, added: true }; });
      });
    });
  }

  g.Anexa = {
    DEFAULTS: DEFAULTS, op: op,
    loadSettings: loadSettings, saveSettings: saveSettings,
    getFile: getFile, setFile: setFile, clearFile: clearFile,
    log: log, getLog: getLog, clearLog: clearLog,
    isValidEmail: isValidEmail, canSend: canSend,
    hasAttachment: hasAttachment, attachDefault: attachDefault
  };
})(typeof window !== "undefined" ? window : this);
