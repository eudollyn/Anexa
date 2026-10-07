/* Anexa - eventos automáticos do Outlook (rodam sem o painel aberto). */
(function () {
  "use strict";

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function done(event) {
    try { event.completed(); } catch (e) { }
  }

  function ready() {
    return typeof Anexa !== "undefined";
  }

  // Lista de destinatários (Para, Cc, Cco) como texto, para saber se algo mudou.
  function readRecipients(item) {
    var A = Anexa;
    return Promise.all([A.op(item.to, "getAsync"), A.op(item.cc, "getAsync"), A.op(item.bcc, "getAsync")])
      .then(function (parts) {
        var all = [].concat(parts[0] || [], parts[1] || [], parts[2] || []);
        var addresses = all.map(function (r) { return String(r.emailAddress || "").trim().toLowerCase(); });
        return { list: addresses, sig: addresses.slice().sort().join(";") };
      });
  }

  function composeType(item) {
    if (typeof item.getComposeTypeAsync !== "function") return Promise.resolve("newMail");
    return Anexa.op(item, "getComposeTypeAsync")
      .then(function (v) { return (v && v.composeType) || "newMail"; })
      .catch(function () { return "newMail"; });
  }

  function sessionGet(item, key) {
    if (!item.sessionData) return Promise.resolve(null);
    return Anexa.op(item.sessionData, "getAsync", key).catch(function () { return null; });
  }
  function sessionSet(item, key, value) {
    if (!item.sessionData) return Promise.resolve();
    return Anexa.op(item.sessionData, "setAsync", key, value).catch(function () { });
  }

  // Envia se houver destinatário válido e ele ficar sem mudanças durante o tempo de espera.
  function trySend(item, s) {
    var A = Anexa;
    return composeType(item).then(function (type) {
      if (type !== "newMail" && !s.sendReplies) return null;
      return readRecipients(item);
    }).then(function (r1) {
      if (!r1 || r1.list.length === 0) return;
      if (s.requireValid && !r1.list.every(A.isValidEmail)) {
        return sessionGet(item, "anexaInvalidLogged").then(function (v) {
          if (v) return;
          A.log("Destinatário inválido. Envio mantido em espera até corrigir.", false);
          return sessionSet(item, "anexaInvalidLogged", "1");
        });
      }
      return sleep(s.delayMs).then(function () { return readRecipients(item); }).then(function (r2) {
        // Mudou durante a espera: o evento disparado pela mudança cuida do envio.
        if (r2.sig !== r1.sig) return;
        return sessionGet(item, "anexaSent").then(function (sent) {
          if (sent) return;
          var check = Promise.resolve(true);
          if (s.validateFile && s.autoAttach && s.fileName) check = A.hasAttachment(item, s.fileName);
          return check.then(function (ok) {
            if (!ok) {
              A.log("Envio cancelado: o anexo " + s.fileName + " não está no e-mail.", false);
              return sessionSet(item, "anexaSent", "cancelado");
            }
            if (!A.canSend()) {
              A.log("Esta versão do Outlook não permite envio por suplemento. O arquivo foi anexado; envie manualmente.", false);
              return sessionSet(item, "anexaSent", "sem-suporte");
            }
            return sessionSet(item, "anexaSent", "1")
              .then(function () { return A.op(item, "saveAsync").catch(function () { }); })
              .then(function () { return A.op(item, "sendAsync"); })
              .then(function () { A.log("E-mail enviado para " + r2.list.join(", ") + ".", true); });
          });
        });
      });
    });
  }

  function onNewMessageComposeHandler(event) {
    if (!ready()) { done(event); return; }
    var A = Anexa;
    var item = Office.context.mailbox.item;
    var s = A.loadSettings();
    if (!s.enabled) { done(event); return; }

    var chain = Promise.resolve();
    if (s.autoAttach) {
      chain = chain.then(function () { return A.attachDefault(item); }).then(function (res) {
        if (res.added) A.log("Arquivo anexado: " + res.name, true);
      });
    }
    chain
      .then(function () { if (s.autoSend) return trySend(item, s); })
      .catch(function (e) { A.log("Não foi possível concluir: " + (e && e.message ? e.message : e), false); })
      .then(function () { done(event); });
  }

  function onMessageRecipientsChangedHandler(event) {
    if (!ready()) { done(event); return; }
    var A = Anexa;
    var item = Office.context.mailbox.item;
    var s = A.loadSettings();
    if (!s.enabled || !s.autoSend) { done(event); return; }
    trySend(item, s)
      .catch(function (e) { A.log("Envio não realizado: " + (e && e.message ? e.message : e), false); })
      .then(function () { done(event); });
  }

  Office.onReady(function () { });
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
  }
  if (typeof window !== "undefined") {
    window.onNewMessageComposeHandler = onNewMessageComposeHandler;
    window.onMessageRecipientsChangedHandler = onMessageRecipientsChangedHandler;
  }
})();
