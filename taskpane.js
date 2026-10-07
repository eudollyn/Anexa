/* Anexa - painel de configuração. */
(function () {
  "use strict";
  var MAX_BYTES = 25 * 1024 * 1024; // limite do Outlook para anexos adicionados por suplemento
  var s;

  function $(id) { return document.getElementById(id); }

  function sizeText(b) {
    if (b >= 1048576) return (b / 1048576).toFixed(1).replace(".", ",") + " MB";
    if (b >= 1024) return Math.round(b / 1024) + " KB";
    return b + " bytes";
  }

  function toast(msg) {
    var t = $("toast");
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast.timer);
    toast.timer = setTimeout(function () { t.hidden = true; }, 3500);
  }

  function save() {
    return Anexa.saveSettings(s).catch(function (e) {
      toast("Não foi possível salvar: " + (e && e.message ? e.message : e));
    });
  }

  function setDot(id, cls) {
    var el = $(id);
    el.className = cls || "";
  }

  function render(file) {
    // Toggles
    var boxes = document.querySelectorAll("input[data-key]");
    for (var i = 0; i < boxes.length; i++) boxes[i].checked = !!s[boxes[i].getAttribute("data-key")];
    document.querySelector('[data-key="sendReplies"]').disabled = !s.autoSend;
    document.querySelector('[data-key="validateFile"]').disabled = !s.autoSend || !s.autoAttach;
    $("delay").value = (s.delayMs / 1000).toString();

    // Arquivo
    var hasFile = !!(file && file.base64);
    if (hasFile) {
      $("fileName").textContent = file.name;
      $("fileMeta").textContent = sizeText(file.size) + " · salvo neste computador";
      $("fileMeta").className = "";
    } else if (s.fileName) {
      $("fileName").textContent = s.fileName;
      $("fileMeta").textContent = "Arquivo não encontrado neste computador. Escolha de novo.";
      $("fileMeta").className = "err";
    } else {
      $("fileName").textContent = "Nenhum arquivo selecionado";
      $("fileMeta").textContent = "Escolha um arquivo para começar";
      $("fileMeta").className = "";
    }

    var item = Office.context.mailbox && Office.context.mailbox.item;
    var composing = !!(item && typeof item.addFileAttachmentFromBase64Async === "function");
    $("attachNow").hidden = !(composing && hasFile);

    // Status
    var supportsEvents = Office.context.requirements.isSetSupported("Mailbox", "1.13");
    var supportsSend = Office.context.requirements.isSetSupported("Mailbox", "1.15");
    if (supportsEvents) { setDot("dotOutlook", "ok"); $("txtOutlook").textContent = "Compatível com a automação"; }
    else { setDot("dotOutlook", "bad"); $("txtOutlook").textContent = "Versão do Outlook sem suporte a eventos automáticos"; }

    if (!s.autoSend) { setDot("dotSend", ""); $("txtSend").textContent = "Desligado: só anexa"; }
    else if (supportsSend) { setDot("dotSend", "ok"); $("txtSend").textContent = "Disponível neste Outlook"; }
    else { setDot("dotSend", "warn-dot"); $("txtSend").textContent = "Não suportado aqui: o arquivo é anexado e você envia"; }

    if (hasFile) { setDot("dotFile", "ok"); $("txtFile").textContent = file.name; }
    else if (s.autoAttach) { setDot("dotFile", "bad"); $("txtFile").textContent = "Escolha o arquivo padrão"; }
    else { setDot("dotFile", ""); $("txtFile").textContent = "Anexo automático desligado"; }

    var pill = $("pill");
    if (!s.enabled) { pill.className = "pill idle"; pill.lastChild.textContent = "Pausado"; }
    else if (!supportsEvents || (s.autoAttach && !hasFile)) { pill.className = "pill warn"; pill.lastChild.textContent = "Atenção"; }
    else { pill.className = "pill on"; pill.lastChild.textContent = "Em execução"; }

    renderLog();
  }

  function renderLog() {
    var ul = $("log");
    var items = Anexa.getLog().slice(0, 30);
    ul.innerHTML = "";
    if (items.length === 0) {
      var li0 = document.createElement("li");
      li0.className = "empty";
      li0.textContent = "Aguardando a primeira ação...";
      ul.appendChild(li0);
      return;
    }
    items.forEach(function (e) {
      var li = document.createElement("li");
      if (!e.ok) li.className = "err";
      var d = new Date(e.t);
      var time = document.createElement("time");
      time.textContent = ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2);
      var dot = document.createElement("i");
      var span = document.createElement("span");
      span.textContent = e.m;
      li.appendChild(dot); li.appendChild(time); li.appendChild(span);
      ul.appendChild(li);
    });
  }

  function refresh() {
    return Anexa.getFile().catch(function () { return null; }).then(render);
  }

  function onFileChosen(ev) {
    var f = ev.target.files && ev.target.files[0];
    ev.target.value = "";
    if (!f) return;
    if (f.size === 0) { toast("Este arquivo está vazio."); return; }
    if (f.size > MAX_BYTES) { toast("O arquivo passa de 25 MB, o limite do Outlook para anexos."); return; }
    var reader = new FileReader();
    reader.onload = function () {
      var data = String(reader.result);
      var base64 = data.substring(data.indexOf(",") + 1);
      Anexa.setFile({ name: f.name, size: f.size, type: f.type, base64: base64, savedAt: Date.now() })
        .then(function () {
          s.fileName = f.name; s.fileSize = f.size;
          Anexa.log("Arquivo padrão definido: " + f.name, true);
          return save();
        })
        .then(refresh)
        .then(function () { toast("Arquivo salvo. Ele será anexado nos próximos e-mails."); })
        .catch(function (e) { toast("Não foi possível salvar o arquivo: " + (e && e.message ? e.message : e)); });
    };
    reader.onerror = function () { toast("Não foi possível ler o arquivo."); };
    reader.readAsDataURL(f);
  }

  Office.onReady(function () {
    s = Anexa.loadSettings();

    var boxes = document.querySelectorAll("input[data-key]");
    for (var i = 0; i < boxes.length; i++) {
      boxes[i].addEventListener("change", function (e) {
        s[e.target.getAttribute("data-key")] = e.target.checked;
        save().then(refresh);
      });
    }
    $("delay").addEventListener("change", function (e) {
      var v = parseFloat(String(e.target.value).replace(",", "."));
      if (isNaN(v)) v = 3;
      s.delayMs = Math.round(Math.min(15, Math.max(1, v)) * 1000);
      save().then(refresh);
    });
    $("fileInput").addEventListener("change", onFileChosen);
    $("attachNow").addEventListener("click", function () {
      Anexa.attachDefault(Office.context.mailbox.item).then(function (r) {
        if (r.added) Anexa.log("Arquivo anexado manualmente: " + r.name, true);
        toast(r.added ? "Arquivo anexado." : "Este e-mail já tem o arquivo.");
        renderLog();
      }).catch(function (e) { toast("Não foi possível anexar: " + (e && e.message ? e.message : e)); });
    });
    $("clearLog").addEventListener("click", function () { Anexa.clearLog(); renderLog(); });

    refresh();
    setInterval(renderLog, 3000);
  });
})();
