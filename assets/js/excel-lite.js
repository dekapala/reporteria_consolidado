(function(){
  const ZIP_SIGNATURE_LOCAL = 0x04034b50;
  const ZIP_SIGNATURE_CENTRAL = 0x02014b50;
  const ZIP_SIGNATURE_END = 0x06054b50;

  const textDecoder = new TextDecoder('utf-8');
  const builtInDateFormats = new Map([
    [14, { hasTime: false }],
    [15, { hasTime: false }],
    [16, { hasTime: false }],
    [17, { hasTime: false }],
    [18, { hasTime: true }],
    [19, { hasTime: true }],
    [20, { hasTime: true }],
    [21, { hasTime: true }],
    [22, { hasTime: true }],
    [27, { hasTime: false }],
    [30, { hasTime: false }],
    [36, { hasTime: false }],
    [45, { hasTime: true }],
    [46, { hasTime: true }],
    [47, { hasTime: true }],
    [50, { hasTime: true }],
    [57, { hasTime: true }]
  ]);

  const looksLikeDateCache = new Map();

  function bytesToString(bytes) {
    if (!bytes) return '';
    try {
      return textDecoder.decode(bytes);
    } catch (err) {
      return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
    }
  }

  function parseXml(bytes) {
    const xml = bytesToString(bytes);
    return new DOMParser().parseFromString(xml, 'application/xml');
  }

  function findEndOfCentralDirectory(data) {
    for (let i = data.length - 22; i >= 0; i--) {
      if (readUInt32LE(data, i) === ZIP_SIGNATURE_END) {
        return i;
      }
    }
    throw new Error('No se encontró el fin del directorio central del ZIP');
  }

  function readUInt16LE(data, offset) {
    return data[offset] | (data[offset + 1] << 8);
  }

  function readUInt32LE(data, offset) {
    return (data[offset]) |
      (data[offset + 1] << 8) |
      (data[offset + 2] << 16) |
      (data[offset + 3] << 24);
  }

  async function inflateRaw(data, expectedSize) {
    if (typeof DecompressionStream === 'function') {
      const stream = new DecompressionStream('deflate-raw');
      const response = new Response(new Blob([data]).stream().pipeThrough(stream));
      const buffer = await response.arrayBuffer();
      const result = new Uint8Array(buffer);
      if (expectedSize && expectedSize !== 0 && result.length !== expectedSize) {
        return result.slice(0, expectedSize);
      }
      return result;
    }
    throw new Error('El navegador no soporta DecompressionStream y no se dispone de un descompresor Deflate embebido.');
  }

  async function unzip(data) {
    const entries = new Map();
    if (!(data instanceof Uint8Array)) {
      data = new Uint8Array(data);
    }

    const eocdOffset = findEndOfCentralDirectory(data);
    const totalEntries = readUInt16LE(data, eocdOffset + 10);
    const centralDirectoryOffset = readUInt32LE(data, eocdOffset + 16);

    let offset = centralDirectoryOffset;
    const decoder = new TextDecoder('utf-8');

    for (let i = 0; i < totalEntries; i++) {
      const signature = readUInt32LE(data, offset);
      if (signature !== ZIP_SIGNATURE_CENTRAL) {
        throw new Error('Firma del directorio central inválida en el ZIP');
      }

      const compression = readUInt16LE(data, offset + 10);
      const compressedSize = readUInt32LE(data, offset + 20);
      const uncompressedSize = readUInt32LE(data, offset + 24);
      const nameLength = readUInt16LE(data, offset + 28);
      const extraLength = readUInt16LE(data, offset + 30);
      const commentLength = readUInt16LE(data, offset + 32);
      const localHeaderOffset = readUInt32LE(data, offset + 42);

      const nameStart = offset + 46;
      const nameBytes = data.subarray(nameStart, nameStart + nameLength);
      const fileName = decoder.decode(nameBytes);

      offset = nameStart + nameLength + extraLength + commentLength;

      const localSignature = readUInt32LE(data, localHeaderOffset);
      if (localSignature !== ZIP_SIGNATURE_LOCAL) {
        throw new Error('Firma local inválida en el ZIP');
      }

      const localNameLength = readUInt16LE(data, localHeaderOffset + 26);
      const localExtraLength = readUInt16LE(data, localHeaderOffset + 28);
      const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
      const compressed = data.subarray(dataStart, dataStart + compressedSize);

      if (compression === 0) {
        entries.set(fileName, compressed.slice());
      } else if (compression === 8) {
        const inflated = await inflateRaw(compressed, uncompressedSize);
        entries.set(fileName, inflated);
      } else {
        console.warn(`Compresión ZIP no soportada para ${fileName} (método ${compression}). Se omite.`);
      }
    }

    return entries;
  }

  function looksLikeDateFormat(formatCode) {
    if (!formatCode) return { isDate: false, hasTime: false };
    if (looksLikeDateCache.has(formatCode)) {
      return looksLikeDateCache.get(formatCode);
    }
    const lower = formatCode.toLowerCase();
    const cleaned = lower
      .replace(/"[^"]*"/g, '')
      .replace(/\[[^\]]*\]/g, '')
      .replace(/\\./g, '');
    const hasDate = /d/.test(cleaned) && /m/.test(cleaned) && /y/.test(cleaned);
    const hasTime = /h/.test(cleaned) || /s/.test(cleaned);
    const result = { isDate: hasDate, hasTime };
    looksLikeDateCache.set(formatCode, result);
    return result;
  }

  function buildStyles(stylesBytes) {
    if (!stylesBytes) {
      return {
        isDateStyle: () => false,
        hasTime: () => false
      };
    }

    const doc = parseXml(stylesBytes);
    const numFmtNodes = doc.querySelectorAll('numFmts numFmt');
    const customFormats = new Map();

    numFmtNodes.forEach(node => {
      const id = parseInt(node.getAttribute('numFmtId') || '0', 10);
      const formatCode = node.getAttribute('formatCode') || '';
      customFormats.set(id, formatCode);
    });

    const styleInfo = [];
    const xfNodes = doc.querySelectorAll('cellXfs xf');

    xfNodes.forEach((xf, idx) => {
      const numFmtId = parseInt(xf.getAttribute('numFmtId') || '0', 10);
      let isDate = false;
      let hasTime = false;

      if (builtInDateFormats.has(numFmtId)) {
        const info = builtInDateFormats.get(numFmtId);
        isDate = true;
        hasTime = Boolean(info && info.hasTime);
      } else if (customFormats.has(numFmtId)) {
        const info = looksLikeDateFormat(customFormats.get(numFmtId));
        isDate = info.isDate;
        hasTime = info.hasTime;
      }

      styleInfo[idx] = {
        isDate,
        hasTime
      };
    });

    return {
      isDateStyle: index => Boolean(styleInfo[index] && styleInfo[index].isDate),
      hasTime: index => Boolean(styleInfo[index] && styleInfo[index].hasTime)
    };
  }

  function parseSharedStrings(sharedBytes) {
    if (!sharedBytes) return [];
    const doc = parseXml(sharedBytes);
    const items = [];
    const siNodes = doc.getElementsByTagName('si');

    for (let i = 0; i < siNodes.length; i++) {
      const node = siNodes[i];
      const textNodes = node.getElementsByTagName('t');
      let text = '';
      for (let j = 0; j < textNodes.length; j++) {
        text += textNodes[j].textContent || '';
      }
      items.push(text);
    }
    return items;
  }

  function columnToIndex(ref) {
    const match = ref.match(/[A-Z]+/i);
    if (!match) return 0;
    const letters = match[0].toUpperCase();
    let index = 0;
    for (let i = 0; i < letters.length; i++) {
      index = index * 26 + (letters.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  function excelSerialToDate(serial) {
    const value = Number(serial);
    if (!isFinite(value)) return null;
    const days = Math.floor(value);
    const millisecondsPerDay = 24 * 60 * 60 * 1000;
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const result = new Date(epoch.getTime() + days * millisecondsPerDay);
    const fractional = value - days;
    if (fractional > 0) {
      const seconds = Math.round(fractional * 24 * 60 * 60);
      result.setUTCSeconds(seconds);
    }
    return result;
  }

  function formatDate(date, includeTime) {
    if (!date) return '';
    const day = String(date.getUTCDate()).padStart(2, '0');
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const year = date.getUTCFullYear();
    if (!includeTime) {
      return `${day}/${month}/${year}`;
    }
    const hours = String(date.getUTCHours()).padStart(2, '0');
    const minutes = String(date.getUTCMinutes()).padStart(2, '0');
    return `${day}/${month}/${year} ${hours}:${minutes}`;
  }

  function parseSheet(sheetBytes, sharedStrings, styles) {
    const doc = parseXml(sheetBytes);
    const rows = [];
    const rowNodes = doc.getElementsByTagName('row');

    for (let i = 0; i < rowNodes.length; i++) {
      const rowNode = rowNodes[i];
      const cells = rowNode.getElementsByTagName('c');
      const row = [];
      let lastCol = -1;

      for (let c = 0; c < cells.length; c++) {
        const cell = cells[c];
        const ref = cell.getAttribute('r');
        let colIndex = ref ? columnToIndex(ref) : lastCol + 1;
        if (colIndex < 0) colIndex = lastCol + 1;

        const type = cell.getAttribute('t');
        const styleIndex = parseInt(cell.getAttribute('s') || '-1', 10);
        const vNode = cell.getElementsByTagName('v')[0];
        const isNode = cell.getElementsByTagName('is')[0];

        let value = '';

        if (type === 's') {
          const idx = vNode ? parseInt(vNode.textContent || '0', 10) : 0;
          value = sharedStrings[idx] || '';
        } else if (type === 'b') {
          value = vNode && vNode.textContent === '1' ? 'TRUE' : 'FALSE';
        } else if (type === 'inlineStr' && isNode) {
          const tNode = isNode.getElementsByTagName('t')[0];
          value = tNode ? tNode.textContent || '' : '';
        } else if (type === 'str') {
          value = vNode ? vNode.textContent || '' : '';
        } else if (type === 'd') {
          const date = vNode ? new Date(vNode.textContent) : null;
          value = date ? formatDate(date, true) : '';
        } else if (vNode) {
          let text = vNode.textContent || '';
          if (styles && styleIndex >= 0 && styles.isDateStyle(styleIndex)) {
            const includeTime = styles.hasTime(styleIndex);
            const parsed = excelSerialToDate(text);
            value = parsed ? formatDate(parsed, includeTime) : text;
          } else {
            value = text;
          }
        } else {
          value = '';
        }

        while (row.length < colIndex) {
          row.push('');
        }
        row[colIndex] = typeof value === 'string' ? value.trim() : value;
        lastCol = colIndex;
      }

      rows.push(row);
    }

    return rows;
  }

  function findHeaderRow(rows) {
    const maxRows = Math.min(rows.length, 25);
    const maxCols = Math.min(rows[0] ? rows[0].length : 0, 20);
    for (let r = 0; r < maxRows; r++) {
      const row = rows[r];
      if (!row) continue;
      for (let c = 0; c < Math.min(row.length, maxCols); c++) {
        const value = String(row[c] || '').toLowerCase();
        if (!value) continue;
        if ((value.includes('zona') && value.includes('tecn')) ||
            (value.includes('numero') && (value.includes('caso') || value.includes('cuenta')))) {
          return r;
        }
      }
    }
    return 0;
  }

  function dedupeHeaders(headers) {
    const seen = new Map();
    return headers.map((header, idx) => {
      let name = header && String(header).trim();
      if (!name) {
        name = `Columna ${idx + 1}`;
      }
      const key = name.toLowerCase();
      if (!seen.has(key)) {
        seen.set(key, 1);
        return name;
      }
      const count = seen.get(key) + 1;
      seen.set(key, count);
      return `${name} (${count})`;
    });
  }

  function rowsToObjects(rows) {
    if (!rows.length) return [];
    const headerIndex = findHeaderRow(rows);
    const headers = dedupeHeaders(rows[headerIndex] || []);
    const result = [];

    for (let r = headerIndex + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      let hasData = false;
      const obj = {};
      for (let c = 0; c < headers.length; c++) {
        const header = headers[c] || `Columna ${c + 1}`;
        let value = row[c];
        if (value === undefined || value === null) value = '';
        if (!hasData && value !== '') {
          hasData = true;
        }
        obj[header] = value;
      }
      if (hasData) {
        result.push(obj);
      }
    }

    return result;
  }

  function resolveSheetTarget(target) {
    if (!target) return 'xl/worksheets/sheet1.xml';
    if (target.startsWith('/')) {
      return target.slice(1);
    }
    if (target.startsWith('xl/')) {
      return target;
    }
    return `xl/${target}`;
  }

  function parseWorkbook(entries) {
    const workbookBytes = entries.get('xl/workbook.xml');
    if (!workbookBytes) {
      throw new Error('El archivo Excel no contiene xl/workbook.xml');
    }
    const workbookDoc = parseXml(workbookBytes);
    const sheetNodes = workbookDoc.getElementsByTagName('sheet');
    const sheets = [];
    for (let i = 0; i < sheetNodes.length; i++) {
      const node = sheetNodes[i];
      sheets.push({
        name: node.getAttribute('name') || `Hoja ${i + 1}`,
        id: node.getAttribute('r:id') || node.getAttribute('id') || `sheet${i + 1}`
      });
    }

    const relsBytes = entries.get('xl/_rels/workbook.xml.rels');
    const relationships = new Map();
    if (relsBytes) {
      const relsDoc = parseXml(relsBytes);
      const relNodes = relsDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < relNodes.length; i++) {
        const rel = relNodes[i];
        const id = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        if (id && target) {
          relationships.set(id, target);
        }
      }
    }

    sheets.forEach((sheet, index) => {
      if (relationships.has(sheet.id)) {
        sheet.target = resolveSheetTarget(relationships.get(sheet.id));
      } else {
        sheet.target = `xl/worksheets/sheet${index + 1}.xml`;
      }
    });

    return sheets;
  }

  async function parseXlsx(data) {
    const entries = await unzip(data);
    const sheets = parseWorkbook(entries);
    if (!sheets.length) {
      throw new Error('El libro Excel no contiene hojas');
    }

    const firstSheet = sheets[0];
    const sheetBytes = entries.get(firstSheet.target);
    if (!sheetBytes) {
      throw new Error(`No se encontró la hoja ${firstSheet.target}`);
    }

    const sharedStrings = parseSharedStrings(entries.get('xl/sharedStrings.xml'));
    const styles = buildStyles(entries.get('xl/styles.xml'));
    const rows = parseSheet(sheetBytes, sharedStrings, styles);
    return rowsToObjects(rows);
  }

  window.ExcelLite = {
    parse: parseXlsx
  };
})();
