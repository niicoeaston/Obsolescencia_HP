/* Cotizaciones de selección. Funciona localmente y no envía datos a servidores.
 * Cargar pdf-lib.min.js y marcas.js antes de este archivo.
 * Los precios de entrada son finales, en CLP y con IVA incluido.
 * Folio, zona, responsable y emisión proceden del registro real del portal.
 * Este componente no crea folios ni mantiene contadores propios.
 */
(function (root) {
  'use strict';
  const PAGE = { width: 595.2756, height: 841.8898, left: 44, right: 551.2756 };
  const COLUMNS = [82, 198, 45, 88, 94.2756];
  const LABELS = ['Código / SAP', 'Descripción', 'Cant.', 'Precio unitario', 'Total'];
  const pesos = n => '$' + n.toLocaleString('es-CL', { maximumFractionDigits: 0 });
  const escapeHTML = value => String(value).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const clean = value => String(value == null ? '' : value).normalize('NFC').trim();
  const VALIDEZ_MAX_MS = 72 * 60 * 60 * 1000;
  const VALIDEZ_TEXTO = 'Válido por un máximo de 3 días desde la emisión.';

  function instante(value, campo) {
    const s = clean(value);
    if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,3})?)?(?:Z|[+-]\d{2}:\d{2})$/.test(s)) throw new Error(campo + ': se requiere fecha y hora ISO con zona horaria.');
    fechaLocal(s.slice(0, 10));
    const valueMs = Date.parse(s);
    if (!Number.isFinite(valueMs)) throw new Error(campo + ' inválida.');
    return valueMs;
  }

  function fechaHoraTexto(value) {
    if (!value) return '____ / ____ / ______  ____:____';
    const parts = new Intl.DateTimeFormat('es-CL', { timeZone:'America/Santiago', year:'numeric', month:'2-digit', day:'2-digit', hour:'2-digit', minute:'2-digit', hourCycle:'h23' }).formatToParts(new Date(value));
    const get = name => parts.find(x => x.type === name).value;
    return get('day') + '/' + get('month') + '/' + get('year') + ' ' + get('hour') + ':' + get('minute');
  }

  function fechaLocal(value) {
    if (value == null || value === '') {
      const d = new Date();
      const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Santiago', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(d);
      const get = name => parts.find(x => x.type === name).value;
      return get('year') + '-' + get('month') + '-' + get('day');
    }
    const s = clean(value);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('La fecha debe usar AAAA-MM-DD.');
    const d = new Date(s + 'T12:00:00Z');
    if (Number.isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== s) throw new Error('Fecha inválida.');
    return s;
  }

  function normalizar(entrada, plantilla) {
    const e = entrada || {};
    const marca = clean(e.marca).toLowerCase();
    if (!root.MarcasCotizacion || !Object.hasOwn(root.MarcasCotizacion, marca)) throw new Error('Marca válida: autoplanet o agroplanet.');
    const lista = e.items == null && plantilla ? [] : e.items;
    if (!Array.isArray(lista)) throw new Error('La selección debe ser una lista de productos.');
    if (!lista.length && !plantilla) throw new Error('Selecciona al menos un producto para cotizar.');
    let total = 0;
    const items = lista.map((item, indice) => {
      const sku = clean(item.sku);
      const descripcion = clean(item.descripcion);
      if (!sku || !descripcion) throw new Error('Falta código o descripción en el producto ' + (indice + 1) + '.');
      if (typeof item.cantidad !== 'number' || !Number.isSafeInteger(item.cantidad) || item.cantidad < 1) throw new Error('La cantidad debe ser un entero mayor a cero en ' + sku + '.');
      if (typeof item.precioUnitario !== 'number' || !Number.isSafeInteger(item.precioUnitario) || item.precioUnitario < 0) throw new Error('El precio debe ser un entero en CLP, con IVA incluido, en ' + sku + '.');
      const importe = item.cantidad * item.precioUnitario;
      if (!Number.isSafeInteger(importe) || !Number.isSafeInteger(total + importe)) throw new Error('El total supera el límite permitido.');
      total += importe;
      return { sku, descripcion, cantidad: item.cantidad, precioUnitario: item.precioUnitario, importe };
    });
    const folio = clean(e.folio), zona = clean(e.zona), responsable = clean(e.responsable);
    if (!plantilla && !zona) throw new Error('Falta la zona de la cotización.');   // AgroPlanet: sin folio y con el responsable en blanco (se completa a mano)
    if (!plantilla && !e.emitidaEn) throw new Error('Falta la fecha y hora de emisión registrada por el portal.');
    let emitidaEn = null, venceEn = null;
    if (e.emitidaEn) {
      const emision = instante(e.emitidaEn, 'Emisión');
      const vencimiento = e.venceEn ? instante(e.venceEn, 'Vencimiento') : emision + VALIDEZ_MAX_MS;
      if (vencimiento <= emision || vencimiento > emision + VALIDEZ_MAX_MS) throw new Error('La validez debe ser positiva y no superar los 3 días (72 horas).');
      emitidaEn = new Date(emision).toISOString();
      venceEn = new Date(vencimiento).toISOString();
    } else if (e.venceEn) throw new Error('El vencimiento requiere una emisión registrada.');
    // A quote represents one zone; do not silently combine stock from another.
    if (zona) for (const item of lista) if (item.zona && clean(item.zona) !== zona) throw new Error('La selección contiene otra zona. Conserva la correspondencia de zona y responsable al agruparla.');
    return { marca, items, total, cliente: clean(e.cliente), folio, zona, responsable, emitidaEn, venceEn, plantilla: !!plantilla };
  }

  const fechaTexto = fecha => fecha ? fecha.split('-').reverse().join('/') : '____ / ____ / ______';

  function renderizar(entrada, opciones) {
    const q = normalizar(entrada, opciones && opciones.plantilla);
    const b = root.MarcasCotizacion[q.marca];
    const filas = q.items.length ? q.items.map(item => '<tr><td>' + escapeHTML(item.sku) + '</td><td>' + escapeHTML(item.descripcion) + '</td><td class="numero">' + item.cantidad + '</td><td class="numero">' + pesos(item.precioUnitario) + '</td><td class="numero">' + pesos(item.importe) + '</td></tr>').join('') : Array.from({ length: 10 }, () => '<tr class="vacia"><td>&nbsp;</td><td></td><td></td><td></td><td></td></tr>').join('');
    return '<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Cotización ' + b.nombre + '</title><style>' +
      ':root{--marca:' + b.color + ';--cabecera:' + b.tintaCabecera + ';--suave:' + b.suave + ';--acento:' + b.acento + ';}*{box-sizing:border-box}body{margin:0;background:#f0f0ee;color:#1d1d1b;font:10pt Arial,Helvetica,sans-serif}.hoja{width:210mm;min-height:297mm;margin:16px auto;background:#fff;padding:19.4mm 15.5mm 14mm}.encabezado{display:flex;align-items:flex-start;justify-content:space-between;gap:10mm;min-height:27mm}.logo{width:65mm;height:auto;max-height:20mm;object-fit:contain;object-position:left top}.identidad{text-align:right;max-width:52%;overflow-wrap:anywhere}h1{font-size:20pt;margin:0 0 5mm;line-height:1.1}.fecha,.folio{font-size:9pt;color:#666;margin:0 0 2mm}.banda{height:1mm;margin:0 0 9mm;background:var(--marca);position:relative}.banda:after{content:"";display:block;position:absolute;right:0;top:0;width:13mm;height:100%;background:var(--acento)}h2{font-size:12pt;margin:0 0 4mm}.moneda{font-size:9pt;color:#666;margin:0 0 7mm}.metadatos{margin:0 0 6mm;line-height:1.55;overflow-wrap:anywhere}.metadatos p{margin:0 0 2mm}.vencimiento{font-size:9pt;color:#666}.cliente{margin:0 0 5mm;overflow-wrap:anywhere}table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:9pt}thead{display:table-header-group}th{text-align:left;background:var(--marca);color:var(--cabecera);font-weight:bold;padding:10px 8px}td{padding:10px 8px;vertical-align:top;overflow-wrap:anywhere;border-bottom:.5px solid #e9e9e7}tbody tr:nth-child(even){background:var(--suave)}tr{break-inside:avoid;page-break-inside:avoid}.numero{text-align:right;font-variant-numeric:tabular-nums}.vacia{height:10.2mm}.resumen{width:86mm;margin:8.5mm 0 0 auto;padding:5mm;background:var(--marca);color:var(--cabecera);break-inside:avoid;page-break-inside:avoid}.resumen span{display:block;font-size:9pt;font-weight:bold}.resumen strong{display:block;text-align:right;font-size:16pt;margin-top:3mm}.pie{border-top:.5px solid #e9e9e7;margin-top:16mm;padding-top:4mm;color:#777;font-size:8pt}@page{size:A4;margin:15mm}@media print{body{background:#fff}.hoja{width:auto;min-height:0;margin:0;padding:0}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}.encabezado{break-inside:avoid}thead th{background:var(--marca)!important}.pie{break-inside:avoid}}@media screen and (max-width:820px){.hoja{margin:0;width:100%;min-height:0;padding:24px}.logo{max-width:42%;height:auto}.encabezado{gap:12px}h1{font-size:18pt}td,th{padding:8px 5px;font-size:8pt}.resumen{max-width:100%}}' +
      '</style></head><body><main class="hoja"><header class="encabezado"><img class="logo" src="' + b.logo + '" alt="' + b.nombre + '"><div class="identidad"><h1>COTIZACIÓN</h1><p class="fecha">Emisión: ' + fechaHoraTexto(q.emitidaEn) + '</p>' + (q.folio || q.plantilla ? '<p class="folio">Folio: ' + escapeHTML(q.folio || '[asignado por el portal]') + '</p>' : '') + '</div></header><div class="banda"></div><h2>Selección de obsolescencias</h2><div class="metadatos"><p><b>Zona:</b> ' + escapeHTML(q.zona || '[zona de los productos seleccionados]') + '</p><p><b>Responsable:</b> ' + (q.responsable ? escapeHTML(q.responsable) : '<span style="display:inline-block;width:70mm;border-bottom:.5px solid #777">&nbsp;</span>') + '</p><p class="vencimiento"><b>Válido hasta:</b> ' + fechaHoraTexto(q.venceEn) + '</p></div><p class="moneda">Precios en pesos chilenos (CLP), IVA incluido.</p>' + (q.cliente ? '<p class="cliente"><b>Cliente:</b> ' + escapeHTML(q.cliente) + '</p>' : '') + '<table><colgroup><col style="width:16.17%"><col style="width:39.03%"><col style="width:8.87%"><col style="width:17.35%"><col style="width:18.58%"></colgroup><thead><tr>' + LABELS.map((s, i) => '<th' + (i >= 2 ? ' class="numero"' : '') + '>' + s + '</th>').join('') + '</tr></thead><tbody>' + filas + '</tbody></table><section class="resumen"><span>TOTAL IVA INCLUIDO</span><strong>' + (q.plantilla && !q.items.length ? '$ ______________' : pesos(q.total)) + '</strong></section><footer class="pie"><strong>' + VALIDEZ_TEXTO + '</strong><br>Portal de obsolescencias</footer></main></body></html>';
  }

  function color(hex) {
    const v = hex.replace('#', '');
    return root.PDFLib.rgb(parseInt(v.slice(0, 2), 16) / 255, parseInt(v.slice(2, 4), 16) / 255, parseInt(v.slice(4, 6), 16) / 255);
  }

  // Unsupported glyphs receive a safe placeholder; Spanish characters are kept.
  function pdfText(s, font) {
    return Array.from(clean(s)).map(ch => { try { font.encodeText(ch); return ch; } catch (_) { return '?'; } }).join('');
  }

  function wrap(value, font, size, width) {
    const text = pdfText(value, font);
    const lines = [];
    let line = '';
    for (const paragraph of text.split(/\r?\n/)) {
      for (const word of paragraph.split(/\s+/)) {
        if (!word) continue;
        if (font.widthOfTextAtSize(word, size) > width) {
          if (line) { lines.push(line); line = ''; }
          for (const ch of word) {
            if (font.widthOfTextAtSize(line + ch, size) > width) { lines.push(line); line = ''; }
            line += ch;
          }
        } else if (line && font.widthOfTextAtSize(line + ' ' + word, size) > width) {
          lines.push(line); line = word;
        } else line += (line ? ' ' : '') + word;
      }
      if (line) { lines.push(line); line = ''; }
    }
    return lines.length ? lines : [''];
  }

  async function crearPDF(entrada, opciones) {
    if (!root.PDFLib) throw new Error('Falta cargar pdf-lib.min.js.');
    const q = normalizar(entrada, opciones && opciones.plantilla);
    const b = root.MarcasCotizacion[q.marca];
    const pdf = await root.PDFLib.PDFDocument.create();
    pdf.setTitle('Cotización ' + b.nombre + ' - Obsolescencias');
    pdf.setAuthor(b.nombre);
    const font = await pdf.embedFont(root.PDFLib.StandardFonts.Helvetica);
    const bold = await pdf.embedFont(root.PDFLib.StandardFonts.HelveticaBold);
    const logo = await pdf.embedPng(b.logo);
    let page;
    let cursor;
    const ink = color('#1D1D1B');
    const grey = color('#666666');
    const lineColor = color('#E9E9E7');
    const pages = [];
    const put = (text, x, top, size, useBold, tint, alignRight) => {
      const f = useBold ? bold : font;
      const s = pdfText(text, f);
      const width = f.widthOfTextAtSize(s, size);
      page.drawText(s, { x: alignRight ? x-width : x, y: PAGE.height-top, size, font: f, color: tint || ink });
    };
    function head() {
      page = pdf.addPage([PAGE.width, PAGE.height]);
      pages.push(page);
      const w = 184;
      const h = w * b.logoAlto / b.logoAncho;
      page.drawImage(logo, { x: PAGE.left, y: PAGE.height-55-h, width: w, height: h });
      put('COTIZACIÓN', PAGE.right, 69, 20, true, ink, true);
      let topEmision = 91;
      if (q.folio || q.plantilla) {
        const folio = wrap('Folio: ' + (q.folio || '[asignado por el portal]'), font, 8.5, 210);
        if (folio.length > 2) throw new Error('La referencia de cotización es demasiado extensa.');
        folio.forEach((v, i) => put(v, PAGE.right, 91+i*10, 8.5, false, grey, true));
        topEmision = 116;
      }
      put('Emisión: ' + fechaHoraTexto(q.emitidaEn), PAGE.right, topEmision, 8.5, false, grey, true);
      page.drawRectangle({ x: PAGE.left, y: PAGE.height-132, width: PAGE.right-PAGE.left, height: 3, color: color(b.color) });
      page.drawRectangle({ x: PAGE.right-36, y: PAGE.height-132, width: 36, height: 3, color: color(b.acento) });
      put('Selección de obsolescencias', PAGE.left, 160, 12, true);
      cursor = 182;
      for (const [label,value] of [['Zona',q.zona || '[zona de los productos seleccionados]'],['Responsable',q.responsable || (q.plantilla ? '[persona responsable de la zona]' : '')]]) {
        if (label === 'Responsable' && !value) {     // en blanco: se completa a mano
          put('Responsable:', PAGE.left, cursor, 10);
          const x0 = PAGE.left + font.widthOfTextAtSize('Responsable: ', 10);
          page.drawLine({ start: { x: x0, y: PAGE.height-cursor-1.5 }, end: { x: x0+210, y: PAGE.height-cursor-1.5 }, thickness: .6, color: grey });
          cursor += 16;
          continue;
        }
        const lines = wrap(label + ': ' + value, font, 10, PAGE.right-PAGE.left);
        if (lines.length > 3) throw new Error(label + ' demasiado extenso para el formato.');
        lines.forEach((s,i) => put(s,PAGE.left,cursor+i*12,10));
        cursor += lines.length*12+4;
      }
      put('Válido hasta: ' + fechaHoraTexto(q.venceEn), PAGE.left, cursor+2, 9, false, grey);
      cursor += 21;
      put('Precios en pesos chilenos (CLP), IVA incluido.', PAGE.left, cursor, 9, false, grey);
      cursor += 20;
      if (q.cliente) {
        const cliente = wrap('Cliente: ' + q.cliente, font, 9, PAGE.right-PAGE.left);
        if (cliente.length > 3) throw new Error('El nombre de cliente es demasiado extenso.');
        cliente.forEach((s, i) => put(s, PAGE.left, cursor+i*11, 9));
        cursor += cliente.length*11 + 12;
      }
      page.drawRectangle({ x: PAGE.left, y: PAGE.height-cursor-29, width: PAGE.right-PAGE.left, height: 29, color: color(b.color) });
      let x = PAGE.left;
      LABELS.forEach((label, i) => {
        put(label, i>=2 ? x+COLUMNS[i]-9 : x+9, cursor+18, 9, true, color(b.tintaCabecera), i>=2);
        x += COLUMNS[i];
      });
      cursor += 29;
    }
    head();
    const rows = q.items.length ? q.items : Array.from({ length: 10 }, () => null);
    rows.forEach((item, index) => {
      const sku = item ? wrap(item.sku, font, 9, COLUMNS[0]-18) : [''];
      const desc = item ? wrap(item.descripcion, font, 9, COLUMNS[1]-18) : [''];
      const nLines = Math.max(sku.length, desc.length);
      const rowHeight = Math.max(29, 17+nLines*11);
      if (rowHeight > 450) throw new Error('La descripción del producto es demasiado extensa para la cotización.');
      const roomForTotal = index === rows.length-1 ? 69 : 0;
      if (cursor+rowHeight+roomForTotal > 740) head();
      if (index%2 === 1) page.drawRectangle({ x: PAGE.left, y: PAGE.height-cursor-rowHeight, width: PAGE.right-PAGE.left, height: rowHeight, color: color(b.suave) });
      page.drawLine({ start: { x: PAGE.left, y: PAGE.height-cursor-rowHeight }, end: { x: PAGE.right, y: PAGE.height-cursor-rowHeight }, thickness: .4, color: lineColor });
      sku.forEach((s, i) => put(s, PAGE.left+9, cursor+18+i*11, 9));
      desc.forEach((s, i) => put(s, PAGE.left+COLUMNS[0]+9, cursor+18+i*11, 9));
      if (item) {
        const nums = [String(item.cantidad), pesos(item.precioUnitario), pesos(item.importe)];
        let x = PAGE.left+COLUMNS[0]+COLUMNS[1];
        nums.forEach((s, i) => {
          let size = 9;
          while (size>6 && font.widthOfTextAtSize(s, size)>COLUMNS[i+2]-18) size-=.25;
          if (font.widthOfTextAtSize(s,size)>COLUMNS[i+2]-18) throw new Error('Importe demasiado largo para el formato de cotización.');
          put(s, x+COLUMNS[i+2]-9, cursor+18, size, false, ink, true);
          x += COLUMNS[i+2];
        });
      }
      cursor += rowHeight;
    });
    let top = cursor+24;
    if (top+45>760) {
      // A final totals page retains the document header without an empty table.
      head();
      top = cursor+8;
    }
    page.drawRectangle({ x: PAGE.right-245, y: PAGE.height-top-45, width: 245, height: 45, color: color(b.color) });
    put('TOTAL IVA INCLUIDO', PAGE.right-245+13, top+18, 9, true, color(b.tintaCabecera));
    const totalTexto = q.plantilla && !q.items.length ? '$ ______________' : pesos(q.total);
    let totalSize = 16;
    while (totalSize>9 && bold.widthOfTextAtSize(totalTexto,totalSize)>219) totalSize-=.25;
    put(totalTexto, PAGE.right-13, top+34, totalSize, true, color(b.tintaCabecera), true);
    pages.forEach((p, i) => {
      p.drawLine({ start: { x: PAGE.left, y: 69 }, end: { x: PAGE.right, y: 69 }, thickness: .4, color: lineColor });
      p.drawText(VALIDEZ_TEXTO, {x:PAGE.left,y:51,size:9,font:bold,color:ink});
      p.drawText('Portal de obsolescencias', { x: PAGE.left, y: 33, size: 8, font, color: grey });
      const count = (i+1) + ' / ' + pages.length;
      p.drawText(count, { x: PAGE.right-font.widthOfTextAtSize(count,8), y: 33, size: 8, font, color: grey });
    });
    return pdf.save();
  }

  async function descargarPDF(entrada, opciones) {
    const bytes = await crearPDF(entrada, opciones);
    const marca = clean(entrada.marca).toLowerCase();
    const q = normalizar(entrada, opciones && opciones.plantilla);
    const folio = clean(entrada.folio).replace(/[^a-zA-Z0-9_-]/g, '_') ||
      (q.emitidaEn ? (q.zona.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-zA-Z0-9]+/g, '') + '_' + fechaHoraTexto(q.emitidaEn).replace(/\D/g, '').replace(/^(\d{2})(\d{2})(\d{4})(\d{4})$/, '$3$2$1_$4')) : 'FORMATO');
    const blob = new Blob([bytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Cotizacion_' + root.MarcasCotizacion[marca].nombre + '_' + folio + '.pdf';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  root.CotizacionObsolescencias = { normalizar: entrada => normalizar(entrada, false), renderizar, crearPDF, descargarPDF, pesos };
})(globalThis);
