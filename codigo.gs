const SHEET_ID = '14EnMYnyrGo29J9sfSJ1svqzVibBmGzC1cNXn6CVnf50';
const SHEET_NAME = 'log-tarefas';
const HEADERS = [
  'Carimbo',
  'nome-executor',
  'nome-categoria',
  'nome-subcategoria',
  'itens-necessarios',
  'cronometro-minutos'
];

function doGet() {
  try {
    const sheet = getTaskSheet_();
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return jsonResponse_({ success: true, data: [] });
    }

    const headerIndexes = getHeaderIndexes_(values[0]);
    const data = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const category = String(row[headerIndexes['nome-categoria']] || '').trim();
      if (!category) continue;

      const stamp = row[headerIndexes.Carimbo];
      const subcategory = String(row[headerIndexes['nome-subcategoria']] || category || '').trim();
      const minutes = Number(row[headerIndexes['cronometro-minutos']]);

      data.push({
        rowNumber: i + 1,
        Carimbo: stamp instanceof Date ? stamp.toISOString() : (stamp ? String(stamp) : ''),
        'nome-executor': String(row[headerIndexes['nome-executor']] || ''),
        'nome-categoria': category,
        'nome-subcategoria': subcategory,
        'itens-necessarios': String(row[headerIndexes['itens-necessarios']] || ''),
        'cronometro-minutos': isFinite(minutes) && minutes >= 0 ? minutes : 0
      });
    }

    return jsonResponse_({ success: true, data: data });
  } catch (error) {
    return jsonResponse_({ success: false, error: error.message });
  }
}

function doPost(event) {
  try {
    const payload = JSON.parse(event.postData.contents);
    const action = String(payload.action || 'create').toLowerCase();

    if (action === 'update') {
      return handleUpdate_(payload);
    }
    return handleCreate_(payload);
  } catch (error) {
    return jsonResponse_({ success: false, error: error.message });
  }
}

function handleCreate_(payload) {
  const executor = String(payload['nome-executor'] || '').trim();
  const category = String(payload['nome-categoria'] || '').trim();
  const subcategory = String(payload['nome-subcategoria'] || '').trim();
  const items = String(payload['itens-necessarios'] || '').trim();
  const minutes = Number(payload['cronometro-minutos']);

  if (!executor || !category) {
    throw new Error('Executor e categoria são obrigatórios.');
  }
  if (!isFinite(minutes) || minutes < 0) {
    throw new Error('O tempo deve ser um número válido maior ou igual a zero.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getTaskSheet_();
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerIndexes = getHeaderIndexes_(headerRow);

    const rowData = new Array(headerRow.length).fill('');
    const dateValue = payload.Carimbo ? new Date(payload.Carimbo) : new Date();

    rowData[headerIndexes.Carimbo] = isNaN(dateValue.getTime()) ? new Date() : dateValue;
    rowData[headerIndexes['nome-executor']] = executor;
    rowData[headerIndexes['nome-categoria']] = category;
    rowData[headerIndexes['nome-subcategoria']] = subcategory;
    rowData[headerIndexes['itens-necessarios']] = items;
    rowData[headerIndexes['cronometro-minutos']] = minutes;

    sheet.appendRow(rowData);
  } finally {
    lock.releaseLock();
  }

  return jsonResponse_({ success: true });
}

function handleUpdate_(payload) {
  const rowNumber = Number(payload.rowNumber);
  if (!Number.isInteger(rowNumber) || rowNumber < 2) {
    throw new Error('rowNumber inválido para atualização.');
  }

  const executor = String(payload['nome-executor'] || '').trim();
  const category = String(payload['nome-categoria'] || '').trim();
  const subcategory = String(payload['nome-subcategoria'] || '').trim();
  const items = String(payload['itens-necessarios'] || '').trim();
  const minutes = Number(payload['cronometro-minutos']);

  if (!executor || !category) {
    throw new Error('Executor e categoria são obrigatórios.');
  }
  if (!isFinite(minutes) || minutes < 0) {
    throw new Error('O tempo deve ser um número válido maior ou igual a zero.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getTaskSheet_();
    const lastRow = sheet.getLastRow();
    if (rowNumber > lastRow) {
      throw new Error('Linha ' + rowNumber + ' não encontrada na planilha.');
    }

    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerIndexes = getHeaderIndexes_(headerRow);

    let dateValue = null;
    if (payload.Carimbo) {
      dateValue = new Date(payload.Carimbo);
      if (isNaN(dateValue.getTime())) {
        dateValue = String(payload.Carimbo);
      }
    }

    if (dateValue !== null) {
      sheet.getRange(rowNumber, headerIndexes.Carimbo + 1).setValue(dateValue);
    }
    sheet.getRange(rowNumber, headerIndexes['nome-executor'] + 1).setValue(executor);
    sheet.getRange(rowNumber, headerIndexes['nome-categoria'] + 1).setValue(category);
    sheet.getRange(rowNumber, headerIndexes['nome-subcategoria'] + 1).setValue(subcategory);
    sheet.getRange(rowNumber, headerIndexes['itens-necessarios'] + 1).setValue(items);
    sheet.getRange(rowNumber, headerIndexes['cronometro-minutos'] + 1).setValue(minutes);
  } finally {
    lock.releaseLock();
  }

  return jsonResponse_({ success: true });
}

function getTaskSheet_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('A aba "' + SHEET_NAME + '" não foi encontrada.');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    getHeaderIndexes_(sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]);
  }
  return sheet;
}

function getHeaderIndexes_(headerRow) {
  const indexes = {};
  HEADERS.forEach(function (header) {
    indexes[header] = headerRow.indexOf(header);
    if (indexes[header] === -1) {
      throw new Error('Cabeçalho obrigatório ausente: ' + header);
    }
  });
  return indexes;
}

function jsonResponse_(body) {
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}
