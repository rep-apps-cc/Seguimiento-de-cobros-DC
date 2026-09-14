/* Seguimiento de Cobros — NegoFIN S.A.E.C.A.
 * Toda la lógica corre en el navegador: no se sube ningún dato a un servidor.
 */
(function () {
  'use strict';

  // ---------------------------------------------------------------
  // Estado
  // ---------------------------------------------------------------
  var state = {
    fileName: '',
    headers: [],          // encabezados originales del excel
    rows: [],             // filas crudas (array de objetos header -> valor)
    numericCols: {},       // header -> true si la columna es numérica
    phoneHeader: null,
    contacts: [],          // contactos procesados (después de "Generar enlaces")
    sentSet: new Set(),    // claves de contactos ya marcados como enviados
    fileKey: ''
  };

  // ---------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------
  function stripAccents(s) {
    return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  function cleanStr(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\u00a0/g, ' ').trim();
  }

  function toToken(header) {
    var t = stripAccents(header).toUpperCase().trim();
    t = t.replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    return t || 'CAMPO';
  }

  function isNumericValue(v) {
    if (typeof v === 'number') return true;
    var s = cleanStr(v).replace(/\./g, '').replace(/,/g, '.');
    return s !== '' && !isNaN(s);
  }

  function toNumber(v) {
    if (typeof v === 'number') return v;
    var s = cleanStr(v).replace(/[^\d.,-]/g, '');
    if (s.indexOf(',') > -1 && s.indexOf('.') > -1) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else if (s.indexOf(',') > -1) {
      s = s.replace(',', '.');
    }
    var n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function formatGs(v) {
    var n = Math.round(toNumber(v));
    return n.toLocaleString('es-PY').replace(/,/g, '.');
  }

  function normalizePhonePY(v) {
    var digits = cleanStr(v).replace(/\D/g, '');
    if (!digits) return '';
    if (digits.startsWith('595')) {
      // ya viene con código de país
    } else if (digits.startsWith('0')) {
      digits = '595' + digits.substring(1);
    } else {
      digits = '595' + digits;
    }
    return digits;
  }

  function isValidPYMobile(digits) {
    // 595 + 9 dígitos (celular paraguayo típico: 595 9XX XXXXXX)
    return /^5959\d{8}$/.test(digits);
  }

  function headerMatches(header, regex) {
    return regex.test(stripAccents(header).toLowerCase());
  }

  function findHeader(headers, regex) {
    for (var i = 0; i < headers.length; i++) {
      if (headerMatches(headers[i], regex)) return headers[i];
    }
    return null;
  }

  function showToast(msg, ms) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.add('is-visible');
    clearTimeout(showToast._tid);
    showToast._tid = setTimeout(function () {
      t.classList.remove('is-visible');
    }, ms || 2200);
  }

  function debounce(fn, wait) {
    var tid;
    return function () {
      clearTimeout(tid);
      var args = arguments;
      tid = setTimeout(function () { fn.apply(null, args); }, wait);
    };
  }

  // ---------------------------------------------------------------
  // Elementos del DOM
  // ---------------------------------------------------------------
  var el = {
    dropZone: document.getElementById('dropZone'),
    fileInput: document.getElementById('fileInput'),
    fileChip: document.getElementById('fileChip'),
    fileChipName: document.getElementById('fileChipName'),
    fileChipClear: document.getElementById('fileChipClear'),
    phonePicker: document.getElementById('phonePicker'),
    phoneColumnSelect: document.getElementById('phoneColumnSelect'),

    statsSection: document.getElementById('statsSection'),
    statTotal: document.getElementById('statTotal'),
    statValid: document.getElementById('statValid'),
    statAmount: document.getElementById('statAmount'),

    templateSection: document.getElementById('templateSection'),
    templateBox: document.getElementById('templateBox'),
    templateToggle: document.getElementById('templateToggle'),
    templateInput: document.getElementById('templateInput'),
    chipsRow: document.getElementById('chipsRow'),
    templatePreviewCount: document.getElementById('templatePreviewCount'),

    searchSection: document.getElementById('searchSection'),
    searchInput: document.getElementById('searchInput'),

    listSection: document.getElementById('listSection'),
    listCount: document.getElementById('listCount'),
    contactList: document.getElementById('contactList'),

    emptyState: document.getElementById('emptyState'),
    actionBar: document.getElementById('actionBar'),
    generateBtn: document.getElementById('generateBtn'),
    exportBtn: document.getElementById('exportBtn'),
    resetBtn: document.getElementById('resetBtn')
  };

  var WHATSAPP_ICON =
    '<svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">' +
    '<path d="M17.5 14.4c-.3-.1-1.6-.8-1.9-.9-.2-.1-.4-.1-.6.1-.2.2-.6.9-.8 1-.1.2-.3.2-.5.1-1.5-.7-2.5-1.3-3.5-2.9-.1-.2-.1-.4.1-.5.2-.2.4-.5.6-.7.1-.2.1-.4 0-.6-.1-.2-.6-1.5-.8-2-.2-.5-.4-.4-.6-.4-.2 0-.4 0-.6 0-.2 0-.5.1-.8.4-.3.3-1 1-1 2.4s1 2.8 1.1 3c.1.2 1.7 2.7 4.2 3.7 2 .8 2.4.6 2.8.6.5-.1 1.6-.7 1.8-1.3.2-.6.2-1.2.2-1.3 0-.1-.1-.2-.3-.3z"/>' +
    '<path d="M12 2C6.5 2 2 6.5 2 12c0 1.9.5 3.6 1.4 5.1L2 22l5-1.3c1.4.8 3.1 1.2 5 1.2 5.5 0 10-4.5 10-10S17.5 2 12 2zm0 18.2c-1.7 0-3.3-.5-4.6-1.3l-.3-.2-3 .8.8-2.9-.2-.3C3.8 15 3.3 13.5 3.3 12 3.3 7.2 7.2 3.3 12 3.3S20.7 7.2 20.7 12 16.8 20.2 12 20.2z"/>' +
    '</svg>';

  // ---------------------------------------------------------------
  // 1. Carga y parseo del Excel
  // ---------------------------------------------------------------
  ['dragover', 'dragenter'].forEach(function (evt) {
    el.dropZone.addEventListener(evt, function (e) {
      e.preventDefault();
      el.dropZone.classList.add('is-drag');
    });
  });
  ['dragleave', 'drop'].forEach(function (evt) {
    el.dropZone.addEventListener(evt, function (e) {
      e.preventDefault();
      el.dropZone.classList.remove('is-drag');
    });
  });
  el.dropZone.addEventListener('drop', function (e) {
    var f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  el.fileInput.addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    if (f) handleFile(f);
  });

  el.fileChipClear.addEventListener('click', function (e) {
    e.preventDefault();
    resetApp();
  });
  el.resetBtn.addEventListener('click', resetApp);

  function handleFile(file) {
    state.fileName = file.name;
    state.fileKey = file.name + '_' + file.size;
    el.fileChipName.textContent = file.name;
    el.fileChip.classList.remove('hidden');
    el.fileInput.value = '';

    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var wb = XLSX.read(data, { type: 'array', cellDates: false });
        var sheetName = wb.SheetNames[0];
        var ws = wb.Sheets[sheetName];
        var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
        processSheet(rows);
      } catch (err) {
        console.error(err);
        showToast('No se pudo leer el archivo. Verificá que sea un Excel válido.', 3200);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function processSheet(rows) {
    if (!rows || !rows.length) {
      showToast('El archivo está vacío.', 3000);
      return;
    }
    // encuentra la primera fila con contenido como encabezado
    var headerRowIdx = 0;
    while (headerRowIdx < rows.length && rows[headerRowIdx].every(function (c) { return cleanStr(c) === ''; })) {
      headerRowIdx++;
    }
    if (headerRowIdx >= rows.length) {
      showToast('No se encontraron datos en el archivo.', 3000);
      return;
    }
    var rawHeaders = rows[headerRowIdx] || [];
    var headers = [];
    var seen = {};
    rawHeaders.forEach(function (h, i) {
      var name = cleanStr(h) || ('Columna ' + (i + 1));
      if (seen[name]) { name = name + ' ' + (++seen[name]); } else { seen[name] = 1; }
      headers.push(name);
    });

    var dataRows = [];
    for (var r = headerRowIdx + 1; r < rows.length; r++) {
      var raw = rows[r];
      if (!raw || raw.every(function (c) { return cleanStr(c) === ''; })) continue;
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = raw[i]; });
      dataRows.push(obj);
    }

    if (!dataRows.length) {
      showToast('El archivo no tiene filas con datos.', 3000);
      return;
    }

    // detectar columnas de MONTO (numéricas y con nombre de monto/deuda) para
    // formatearlas con separador de miles. El resto de columnas numéricas
    // (cédula, teléfono, códigos, etc.) se muestran tal cual, sin reformatear.
    var AMOUNT_HEADER_RE = /monto|cuota|total|deuda|saldo|\bcanc\b|precio|importe|\bpago\b|valor|descuento/;
    var numericCols = {};
    headers.forEach(function (h) {
      var sample = 0, numeric = 0;
      for (var i = 0; i < Math.min(dataRows.length, 25); i++) {
        var v = dataRows[i][h];
        if (cleanStr(v) === '') continue;
        sample++;
        if (isNumericValue(v)) numeric++;
      }
      var mostlyNumeric = sample > 0 && numeric / sample > 0.8;
      numericCols[h] = mostlyNumeric && headerMatches(h, AMOUNT_HEADER_RE);
    });

    state.headers = headers;
    state.rows = dataRows;
    state.numericCols = numericCols;

    detectPhoneColumn();
  }

  // ---------------------------------------------------------------
  // 2. Detección de columnas (teléfono + roles para la plantilla)
  // ---------------------------------------------------------------
  function detectPhoneColumn() {
    var guess = findHeader(state.headers, /cel|whats|movil|tel[ei]?fono|\bnro\b|numero/);
    if (guess) {
      state.phoneHeader = guess;
      el.phonePicker.classList.add('hidden');
      onColumnsReady();
      return;
    }
    // no se pudo adivinar: preguntar al usuario
    el.phoneColumnSelect.innerHTML = state.headers.map(function (h) {
      return '<option value="' + h.replace(/"/g, '&quot;') + '">' + h + '</option>';
    }).join('');
    el.phonePicker.classList.remove('hidden');
    state.phoneHeader = state.headers[0];
    onColumnsReady();
  }

  el.phoneColumnSelect.addEventListener('change', function () {
    state.phoneHeader = el.phoneColumnSelect.value;
  });

  function onColumnsReady() {
    buildChips();
    buildDefaultTemplate();
    renderStats();
    el.statsSection.classList.remove('hidden');
    el.templateSection.classList.remove('hidden');
    el.emptyState.classList.add('hidden');
    el.actionBar.classList.add('is-visible');
    el.exportBtn.disabled = true;
    el.listSection.classList.add('hidden');
    el.searchSection.classList.add('hidden');
    el.contactList.innerHTML = '';
    state.contacts = [];
  }

  // ---------------------------------------------------------------
  // 3. Plantilla del mensaje
  // ---------------------------------------------------------------
  function buildChips() {
    el.chipsRow.innerHTML = state.headers.map(function (h) {
      return '<button type="button" class="chip" data-token="{{' + toToken(h) + '}}">' + h + '</button>';
    }).join('');
    Array.prototype.forEach.call(el.chipsRow.querySelectorAll('.chip'), function (btn) {
      btn.addEventListener('click', function () {
        insertAtCursor(el.templateInput, btn.getAttribute('data-token'));
      });
    });
  }

  function insertAtCursor(textarea, text) {
    var start = textarea.selectionStart || textarea.value.length;
    var end = textarea.selectionEnd || textarea.value.length;
    var val = textarea.value;
    textarea.value = val.slice(0, start) + text + val.slice(end);
    var pos = start + text.length;
    textarea.focus();
    textarea.setSelectionRange(pos, pos);
  }

  el.templateToggle.addEventListener('click', function () {
    el.templateBox.classList.toggle('is-open');
  });

  function buildDefaultTemplate() {
    var headers = state.headers;
    var H = function (regex) { return findHeader(headers, regex); };

    var nombre = H(/nombre|cliente/);
    var cedula = H(/cedula|documento|\bci\b/);
    var marca = H(/marca|empresa/);
    var cuota = H(/cuota/);
    var total = H(/total.*deuda|deuda.*total|\bsaldo\b|deuda/);
    var canc = H(/canc|descuento/);

    var lines = [];
    lines.push('Sr/a' + (nombre ? ' {{' + toToken(nombre) + '}}' : '') + (cedula ? ', CI {{' + toToken(cedula) + '}}' : '') + '.');

    if (marca) {
      lines.push('{{' + toToken(marca) + '}} le brinda estas 2 opciones para deslindarse del proceso judicial en su contra:');
    } else {
      lines.push('Le brindamos estas 2 opciones para deslindarse del proceso judicial en su contra:');
    }

    if (cuota) {
      lines.push('1) Abonar el pago minimo de 1 cuota de Gs. {{' + toToken(cuota) + '}} (monto conversable).');
    }
    if (total && canc) {
      lines.push('2) Cancelar definitivamente su deuda y retirar su libre deuda con un descuento del 25%, en vez de pagar Gs. {{' + toToken(total) + '}} solo pagara Gs. {{' + toToken(canc) + '}}.');
    } else if (total) {
      lines.push('2) Cancelar definitivamente su deuda de Gs. {{' + toToken(total) + '}} y retirar su libre deuda.');
    }

    if (lines.length <= 2) {
      // no se detectaron columnas de monto: plantilla genérica con todos los campos disponibles
      lines = ['Sr/a' + (nombre ? ' {{' + toToken(nombre) + '}}' : '') + ', le contactamos para coordinar el pago de su deuda pendiente.'];
      headers.forEach(function (h) {
        if (h === state.phoneHeader) return;
        lines.push(h + ': {{' + toToken(h) + '}}');
      });
    }

    el.templateInput.value = lines.join('\n');
  }

  // ---------------------------------------------------------------
  // 4. Generar enlaces de WhatsApp
  // ---------------------------------------------------------------
  el.generateBtn.addEventListener('click', function () {
    generateContacts();
  });

  function fillTemplate(template, row) {
    return state.headers.reduce(function (msg, h) {
      var token = '{{' + toToken(h) + '}}';
      var raw = row[h];
      var value = state.numericCols[h] ? formatGs(raw) : cleanStr(raw);
      return msg.split(token).join(value);
    }, template);
  }

  function contactKey(row, idx) {
    return state.fileKey + '::' + idx + '::' + cleanStr(row[state.phoneHeader]);
  }

  function guessNameHeader() {
    return findHeader(state.headers, /nombre|cliente/) || state.headers[0];
  }
  function guessAmountHeader() {
    return findHeader(state.headers, /total.*deuda|deuda.*total|\bsaldo\b|deuda|monto/);
  }

  function generateContacts() {
    var template = el.templateInput.value;
    var nameHeader = guessNameHeader();
    var amountHeader = guessAmountHeader();
    var sentKey = 'cobros_sent_' + state.fileKey;
    var storedSent = [];
    try { storedSent = JSON.parse(localStorage.getItem(sentKey) || '[]'); } catch (e) {}
    state.sentSet = new Set(storedSent);

    state.contacts = state.rows.map(function (row, idx) {
      var digits = normalizePhonePY(row[state.phoneHeader]);
      var valid = isValidPYMobile(digits);
      var message = fillTemplate(template, row);
      var key = contactKey(row, idx);
      return {
        key: key,
        name: cleanStr(row[nameHeader]) || 'Sin nombre',
        cedulaOrPhone: cleanStr(row[state.phoneHeader]),
        amount: amountHeader ? formatGs(row[amountHeader]) : '',
        valid: valid,
        phone: digits,
        message: message,
        link: valid ? ('https://wa.me/' + digits + '?text=' + encodeURIComponent(message)) : null,
        sent: state.sentSet.has(key)
      };
    });

    el.exportBtn.disabled = false;
    renderStats();
    renderList();
    el.searchSection.classList.remove('hidden');
    el.listSection.classList.remove('hidden');
    showToast('Se generaron ' + state.contacts.filter(function (c) { return c.valid; }).length + ' enlaces de WhatsApp.');
  }

  // ---------------------------------------------------------------
  // 5. Resumen / estadísticas
  // ---------------------------------------------------------------
  function renderStats() {
    var total = state.rows.length;
    var amountHeader = guessAmountHeader();
    var sumAmount = 0;
    if (amountHeader) {
      state.rows.forEach(function (row) { sumAmount += toNumber(row[amountHeader]); });
    }
    var validCount;
    if (state.contacts.length) {
      validCount = state.contacts.filter(function (c) { return c.valid; }).length;
    } else {
      validCount = state.rows.filter(function (row) { return isValidPYMobile(normalizePhonePY(row[state.phoneHeader])); }).length;
    }
    el.statTotal.textContent = total.toLocaleString('es-PY');
    el.statValid.textContent = validCount.toLocaleString('es-PY');
    el.statAmount.textContent = amountHeader ? ('Gs. ' + formatGs(sumAmount)) : '—';
  }

  // ---------------------------------------------------------------
  // 6. Lista de contactos
  // ---------------------------------------------------------------
  el.searchInput.addEventListener('input', debounce(function () { renderList(); }, 120));

  function renderList() {
    var q = stripAccents(el.searchInput.value || '').toLowerCase().trim();
    var items = state.contacts.filter(function (c) {
      if (!q) return true;
      return stripAccents(c.name).toLowerCase().indexOf(q) > -1 ||
             stripAccents(c.cedulaOrPhone).toLowerCase().indexOf(q) > -1 ||
             c.phone.indexOf(q) > -1;
    });

    el.listCount.textContent = items.length + ' de ' + state.contacts.length;

    if (!items.length) {
      el.contactList.innerHTML = '<div class="empty"><span class="empty__icon">🔎</span><div class="empty__title">Sin resultados</div><div class="empty__desc">Probá con otro nombre, cédula o número.</div></div>';
      return;
    }

    el.contactList.innerHTML = items.map(function (c, i) {
      return (
        '<div class="row' + (c.valid ? '' : ' is-invalid') + '" data-key="' + c.key + '">' +
          '<div class="row__main">' +
            '<div class="row__name">' + escapeHtml(c.name) + '</div>' +
            '<div class="row__meta">' + escapeHtml(c.cedulaOrPhone || '') + '</div>' +
          '</div>' +
          (c.amount ? '<div class="row__amount">Gs. ' + c.amount + '</div>' : '') +
          (c.valid ?
            '<button class="row__send' + (c.sent ? ' is-sent' : '') + '" data-action="send" aria-label="Enviar WhatsApp">' + (c.sent ? '&#10003;' : WHATSAPP_ICON) + '</button>'
            : '<span class="row__badge">N&uacute;mero inv&aacute;lido</span>'
          ) +
        '</div>'
      );
    }).join('');

    Array.prototype.forEach.call(el.contactList.querySelectorAll('[data-action="send"]'), function (btn) {
      btn.addEventListener('click', function () {
        var rowEl = btn.closest('.row');
        var key = rowEl.getAttribute('data-key');
        var contact = state.contacts.find(function (c) { return c.key === key; });
        if (!contact || !contact.link) return;
        window.open(contact.link, '_blank');
        contact.sent = true;
        state.sentSet.add(key);
        persistSent();
        btn.classList.add('is-sent');
        btn.innerHTML = '&#10003;';
      });
    });
  }

  function persistSent() {
    var sentKey = 'cobros_sent_' + state.fileKey;
    try { localStorage.setItem(sentKey, JSON.stringify(Array.from(state.sentSet))); } catch (e) {}
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  // ---------------------------------------------------------------
  // 7. Exportar Excel con los enlaces
  // ---------------------------------------------------------------
  el.exportBtn.addEventListener('click', function () {
    if (!state.contacts.length) return;
    var aoa = [state.headers.concat(['LINK WHATSAPP'])];
    state.rows.forEach(function (row, idx) {
      var contact = state.contacts[idx];
      var line = state.headers.map(function (h) { return row[h]; });
      line.push(contact && contact.valid ? contact.link : 'Numero invalido');
      aoa.push(line);
    });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    // convertir la última columna en hipervínculos reales
    for (var r = 1; r < aoa.length; r++) {
      var contact = state.contacts[r - 1];
      if (contact && contact.valid) {
        var cellRef = XLSX.utils.encode_cell({ r: r, c: state.headers.length });
        ws[cellRef].l = { Target: contact.link, Tooltip: 'Enviar WhatsApp' };
        ws[cellRef].v = 'Enviar WhatsApp';
      }
    }
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Cobros');
    var outName = state.fileName.replace(/\.(xlsx|xls|csv)$/i, '') + '_whatsapp.xlsx';
    XLSX.writeFile(wb, outName);
    showToast('Excel descargado: ' + outName);
  });

  // ---------------------------------------------------------------
  // 8. Reiniciar
  // ---------------------------------------------------------------
  function resetApp() {
    state = {
      fileName: '', headers: [], rows: [], numericCols: {}, phoneHeader: null,
      contacts: [], sentSet: new Set(), fileKey: ''
    };
    el.fileChip.classList.add('hidden');
    el.phonePicker.classList.add('hidden');
    el.statsSection.classList.add('hidden');
    el.templateSection.classList.add('hidden');
    el.searchSection.classList.add('hidden');
    el.listSection.classList.add('hidden');
    el.emptyState.classList.remove('hidden');
    el.actionBar.classList.remove('is-visible');
    el.contactList.innerHTML = '';
    el.searchInput.value = '';
    el.templateInput.value = '';
    el.exportBtn.disabled = true;
  }

  // ---------------------------------------------------------------
  // 9. Service worker (instalación / pantalla completa offline-ready)
  // ---------------------------------------------------------------
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('service-worker.js').catch(function (err) {
        console.warn('No se pudo registrar el service worker', err);
      });
    });
  }
})();
