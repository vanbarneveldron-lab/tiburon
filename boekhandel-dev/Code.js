
const DB_SCHEMA = {
  Items: [
    'item_id', 'reference_no', 'item_type_id', 'title', 'author',
    'illustrator', 'collection', 'publisher', 'publication_year',
    'isbn', 'cover_type', 'dimensions', 'pages', 'weight',
    'edition', 'condition', 'asking_price', 'lot_label',
    'repetit_flag', 'photo_status', 'status', 'created_at',
    'updated_at'
  ],
  // ItemTypes sheet columns: stable ID, Catalan/Spanish/English labels, then active flag.
  // Keep item_type_id values stable; display labels fall back to name_ca when a translation is blank.
  ItemTypes: ['item_type_id', 'name_ca', 'name_es', 'name_en', 'active'],
  Platforms: ['platform_id', 'name', 'code', 'active'],
  Publications: [
    'publication_id', 'item_id', 'platform_id', 'status',
    'listed_at', 'withdrawn_at'
  ],
  Sales: [
    'sale_id', 'item_id', 'reference_no', 'sale_date',
    'sales_channel', 'gross_price', 'vat_rate', 'vat_amount',
    'net_amount', 'invoice_created', 'invoice_reference', 'status'
  ],
  LotComponents: ['component_id', 'item_id', 'description'],
  Settings: ['setting_key', 'setting_value'],
  ChangeLog: [
    'log_id', 'timestamp', 'user', 'entity',
    'entity_id', 'action', 'details'
  ]
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Velles i Novelles')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Creates missing database tabs without clearing existing data.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.entries(DB_SCHEMA).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold');
    }
  });

  const settings = ss.getSheetByName('Settings');
  const existing = settings.getDataRange().getValues();
  const keys = existing.slice(1).map(row => row[0]);

  const defaults = [
    ['NEXT_REFERENCE_NO', ''],
    ['APP_LANGUAGE', 'ca'],
    ['DATABASE_VERSION', '0.1']
  ];

  defaults.forEach(([key, value]) => {
    if (!keys.includes(key)) {
      settings.appendRow([key, value]);
    }
  });

  return 'Database structure created successfully.';
}

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(name);

  if (!sheet) {
    throw new Error('Database tab not found: ' + name);
  }

  return sheet;
}

function getRecords(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  const headers = values[0];

  return values.slice(1)
    .filter(row => row.some(value => value !== ''))
    .map(row => {
      const record = {};

      headers.forEach((header, index) => {
        const value = row[index];

        // Date objects must be converted before returning
        // data through google.script.run.
        record[header] = value instanceof Date
          ? value.toISOString()
          : value;
      });

      return record;
    });
}

function getSetting(key) {
  const settings = getRecords('Settings');
  const setting = settings.find(row => row.setting_key === key);
  return setting ? setting.setting_value : '';
}

function getInitialData() {
  return {
    items: getRecords('Items'),
    itemTypes: getRecords('ItemTypes'),
    platforms: getRecords('Platforms'),
    sales: getRecords('Sales'),
    nextReference: getSetting('NEXT_REFERENCE_NO'),
    language: getSetting('APP_LANGUAGE') || 'ca'
  };
}

function saveItem(form) {
  if (!form || !String(form.title || '').trim()) {
    throw new Error('Cal introduir un títol o una descripció.');
  }

  if (!form.item_type_id) {
    throw new Error('Cal seleccionar un tipus d’article.');
  }

  const itemTypes = getRecords('ItemTypes');
  const selectedType = itemTypes.find(type =>
    String(type.item_type_id) === String(form.item_type_id) &&
    String(type.active).toUpperCase() !== 'FALSE'
  );

  if (!selectedType) {
    throw new Error('El tipus d’article seleccionat no és vàlid.');
  }

  const price = parsePrice(form.asking_price);

  if (price !== '' && (!Number.isFinite(price) || price < 0)) {
    throw new Error('El preu introduït no és vàlid.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const nextReference = Number(getSetting('NEXT_REFERENCE_NO'));

    if (!Number.isInteger(nextReference) || nextReference < 1) {
      throw new Error(
        'Cal configurar NEXT_REFERENCE_NO al full Settings.'
      );
    }

    const items = getRecords('Items');

    if (items.some(item =>
      Number(item.reference_no) === nextReference
    )) {
      throw new Error('Aquest número de referència ja existeix.');
    }

    const now = new Date();
    const itemId = Utilities.getUuid();

    const item = {
      item_id: itemId,
      reference_no: nextReference,
      item_type_id: form.item_type_id,
      title: String(form.title).trim(),
      author: String(form.author || '').trim(),
      illustrator: String(form.illustrator || '').trim(),
      collection: String(form.collection || '').trim(),
      publisher: String(form.publisher || '').trim(),
      publication_year: form.publication_year || '',
      isbn: String(form.isbn || '').trim(),
      cover_type: String(form.cover_type || '').trim(),
      dimensions: String(form.dimensions || '').trim(),
      pages: form.pages || '',
      weight: form.weight || '',
      edition: String(form.edition || '').trim(),
      condition: String(form.condition || '').trim(),
      asking_price: price,
      lot_label: '',
      repetit_flag: false,
      photo_status: '',
      status: 'active',
      created_at: now,
      updated_at: now
    };

    appendRecord('Items', item);
    setSetting('NEXT_REFERENCE_NO', nextReference + 1);

    logChange(
      'Items',
      itemId,
      'CREATE',
      'Created item ' + nextReference
    );

    return {
      success: true,
      reference_no: nextReference
    };

  } finally {
    lock.releaseLock();
  }
}

function appendRecord(sheetName, record) {
  const sheet = getSheet(sheetName);
  const headers = DB_SCHEMA[sheetName];

  const row = headers.map(header =>
    record[header] !== undefined ? record[header] : ''
  );

  sheet.appendRow(row);
}

function setSetting(key, value) {
  const sheet = getSheet('Settings');
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  sheet.appendRow([key, value]);
}

function logChange(entity, entityId, action, details) {
  appendRecord('ChangeLog', {
    log_id: Utilities.getUuid(),
    timestamp: new Date(),
    user: Session.getActiveUser().getEmail(),
    entity: entity,
    entity_id: entityId,
    action: action,
    details: details
  });
}

function parsePrice(value) {
  if (value === '' || value == null) return '';

  const normalized = String(value).trim().replace(',', '.');

  if (!/^\d+(\.\d{1,2})?$/.test(normalized)) {
    throw new Error(
      'Preu no vàlid. Utilitza, per exemple, 12,50.'
    );
  }

  const price = Number(normalized);

  if (!Number.isFinite(price) || price < 0) {
    throw new Error('El preu no és vàlid.');
  }

  return Math.round(price * 100) / 100;
}

function updateItem(itemId, form) {
  if (!itemId || !form || !String(form.title || '').trim()) {
    throw new Error('Les dades de l’article no són vàlides.');
  }

  if (!form.item_type_id) {
    throw new Error('Cal seleccionar un tipus d’article.');
  }

  const itemTypes = getRecords('ItemTypes');

  const selectedType = itemTypes.find(type =>
    String(type.item_type_id) === String(form.item_type_id) &&
    String(type.active).toUpperCase() !== 'FALSE'
  );

  if (!selectedType) {
    throw new Error('El tipus d’article seleccionat no és vàlid.');
  }

  const price = parsePrice(form.asking_price);

  if (price !== '' && (!Number.isFinite(price) || price < 0)) {
    throw new Error('El preu introduït no és vàlid.');
  }

  const year = form.publication_year === ''
    ? ''
    : Number(form.publication_year);

  if (year !== '' && (!Number.isInteger(year) || year < 0 || year > 2100)) {
    throw new Error('L’any de publicació no és vàlid.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet('Items');
    const values = sheet.getDataRange().getValues();
    const headers = values[0];

    const idColumn = headers.indexOf('item_id');

    if (idColumn < 0) {
      throw new Error('No s’ha trobat la columna item_id.');
    }

    const rowIndex = values.findIndex((row, index) =>
      index > 0 && String(row[idColumn]) === String(itemId)
    );

    if (rowIndex < 1) {
      throw new Error('No s’ha trobat l’article.');
    }

    const currentStatus = values[rowIndex][headers.indexOf('status')];

    if (String(currentStatus).toLowerCase() !== 'active') {
      throw new Error('Només es poden modificar articles actius.');
    }

    const updates = {
      item_type_id: form.item_type_id,
      title: String(form.title).trim(),
      author: String(form.author || '').trim(),
      illustrator: String(form.illustrator || '').trim(),
      collection: String(form.collection || '').trim(),
      publisher: String(form.publisher || '').trim(),
      publication_year: year,
      isbn: String(form.isbn || '').trim(),
      cover_type: String(form.cover_type || '').trim(),
      dimensions: String(form.dimensions || '').trim(),
      pages: form.pages || '',
      weight: form.weight || '',
      edition: String(form.edition || '').trim(),
      condition: String(form.condition || '').trim(),
      asking_price: price,
      updated_at: new Date()
    };

    Object.entries(updates).forEach(([field, value]) => {
      const column = headers.indexOf(field);

      if (column >= 0) {
        sheet.getRange(rowIndex + 1, column + 1).setValue(value);
      }
    });

    const referenceColumn = headers.indexOf('reference_no');
    const reference = values[rowIndex][referenceColumn];

    logChange(
      'Items',
      itemId,
      'UPDATE',
      'Updated item ' + reference
    );

    return {
      success: true,
      reference_no: reference
    };

  } finally {
    lock.releaseLock();
  }
}

function isValidIsbn(isbn) {
  if (/^\d{13}$/.test(isbn)) {
    const sum = isbn.split('').reduce((total, digit, index) => {
      return total + Number(digit) * (index % 2 === 0 ? 1 : 3);
    }, 0);

    return sum % 10 === 0;
  }

  if (/^\d{9}[\dX]$/.test(isbn)) {
    const sum = isbn.split('').reduce((total, char, index) => {
      const value = char === 'X' ? 10 : Number(char);
      return total + value * (10 - index);
    }, 0);

    return sum % 11 === 0;
  }

  return false;
}


function lookupBookByIsbn(isbn) {
  const cleanIsbn = String(isbn || '')
    .replace(/[-\s]/g, '')
    .toUpperCase();

  // Controleer ISBN-nummer
  if (!isValidIsbn(cleanIsbn)) {
    throw new Error('Dit ISBN is ongeldig. Controleer het nummer.');
  }

  // Open Library ISBN API
  const url =
    'https://openlibrary.org/api/books?' +
    'bibkeys=' + encodeURIComponent('ISBN:' + cleanIsbn) +
    '&jscmd=data&format=json';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'User-Agent': 'Boekhandel/0.1 (contact: vanbarneveldron@gmail.com)'
    },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();

  if (status !== 200) {
    Logger.log('Open Library HTTP status: ' + status);
    Logger.log('Open Library response: ' + response.getContentText());

    throw new Error(
      'Open Library geeft HTTP-fout ' + status +
      '. Probeer het later opnieuw.'
    );
  }

  const data = JSON.parse(response.getContentText());
  const book = data['ISBN:' + cleanIsbn];

  if (!book) {
    return {
      found: false,
      message: 'Geen boek gevonden voor dit ISBN.'
    };
  }

  const cover = book.cover || {};

  return {
    found: true,
    source: 'Open Library',
    title: book.title || '',
    author: (book.authors || [])
      .map(author => author.name)
      .join(', '),
    publisher: (book.publishers || [])
      .map(publisher => publisher.name)
      .join(', '),
    publication_year: book.publish_date
      ? (book.publish_date.match(/\d{4}/) || [''])[0]
      : '',
    isbn: cleanIsbn,
    pages: book.number_of_pages || '',
    dimensions: '',
    cover_image: cover.large || cover.medium || cover.small || '',
    source_url: book.url || ''
  };
}
