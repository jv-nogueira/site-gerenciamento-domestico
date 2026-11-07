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
    const data = values.slice(1)
      .filter(function (row) {
        return row[headerIndexes['nome-categoria']] &&
          row[headerIndexes['nome-subcategoria']];
      })
      .map(function (row) {
        const stamp = row[headerIndexes.Carimbo];
        const minutes = Number(row[headerIndexes['cronometro-minutos']]);
        return {
          Carimbo: stamp instanceof Date ? stamp.toISOString() : stamp,
          'nome-executor': String(row[headerIndexes['nome-executor']] || ''),
          'nome-categoria': String(row[headerIndexes['nome-categoria']]),
          'nome-subcategoria': String(row[headerIndexes['nome-subcategoria']]),
          'itens-necessarios': String(row[headerIndexes['itens-necessarios']] || ''),
          'cronometro-minutos': isFinite(minutes) ? minutes : 0
        };
      });

    return jsonResponse_({ success: true, data: data });
  } catch (error) {
    return jsonResponse_({ success: false, error: error.message });
  }
}

function doPost(event) {
  try {
    const payload = JSON.parse(event.postData.contents);
    const executor = String(payload['nome-executor'] || '').trim();
    const category = String(payload['nome-categoria'] || '').trim();
    const subcategory = String(payload['nome-subcategoria'] || '').trim();
    const items = String(payload['itens-necessarios'] || '').trim();
    const minutes = Number(payload['cronometro-minutos']);

    if (!executor || !category || !subcategory) {
      throw new Error('Executor, categoria e tarefa são obrigatórios.');
    }
    if (!isFinite(minutes) || minutes < 0) {
      throw new Error('O tempo deve ser um número válido maior ou igual a zero.');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getTaskSheet_();
      sheet.appendRow([
        payload.Carimbo ? new Date(payload.Carimbo) : new Date(),
        executor,
        category,
        subcategory,
        items,
        minutes
      ]);
    } finally {
      lock.releaseLock();
    }

    return jsonResponse_({ success: true });
  } catch (error) {
    return jsonResponse_({ success: false, error: error.message });
  }
}

function getTaskSheet_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('A aba "' + SHEET_NAME + '" não foi encontrada.');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    getHeaderIndexes_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
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
