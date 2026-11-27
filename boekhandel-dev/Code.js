
const DB_SCHEMA = {
  Items: [
    'item_id', 'reference_no', 'item_type_id', 'title', 'author',
    'illustrator', 'collection', 'publisher', 'publication_year',
    'isbn', 'cover_type', 'dimensions', 'pages', 'weight',
    'edition', 'condition', 'asking_price', 'lot_label',
    'repetit_flag', 'photo_status', 'status', 'created_at',
    'updated_at'
  ],
  // ItemTypes columns match the DEV sheet: stable ID, Catalan label, active flag, then Spanish/English labels.
  // Keep item_type_id values stable; display labels fall back to name_ca when a translation is blank.
  ItemTypes: ['item_type_id', 'name_ca', 'active', 'name_es', 'name_en'],
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
    publications: getRecords('Publications'),
    sales: getRecords('Sales'),
    nextReference: getSetting('NEXT_REFERENCE_NO'),
    language: getSetting('APP_LANGUAGE') || 'ca'
  };
}

function recordPublication(form) {
  if (!form || !String(form.item_id || '').trim()) throw new Error('PUBLICATION_ITEM_REQUIRED');
  const item = getRecords('Items').find(row => String(row.item_id) === String(form.item_id));
  if (!item || String(item.status).toLowerCase() !== 'active') throw new Error('PUBLICATION_ITEM_INVALID');

  const rawPlatformIds = Array.isArray(form.platform_ids)
    ? form.platform_ids
    : [form.platform_id];
  const platformIds = [...new Set(rawPlatformIds.map(id => String(id || '').trim()).filter(Boolean))];
  if (!platformIds.length) throw new Error('PUBLICATION_PLATFORM_REQUIRED');

  const activePlatforms = getRecords('Platforms')
    .filter(row => String(row.active).toUpperCase() !== 'FALSE');
  const platforms = platformIds.map(platformId =>
    activePlatforms.find(row => String(row.platform_id) === platformId)
  );
  if (platforms.some(platform => !platform)) throw new Error('PUBLICATION_PLATFORM_INVALID');

  const listedAt = String(form.listed_at || '').trim();
  if (!isValidSaleDate(listedAt)) throw new Error('PUBLICATION_DATE_INVALID');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    platforms.forEach(platform => {
      const id = Utilities.getUuid();
      appendRecord('Publications', {
        publication_id: id,
        item_id: item.item_id,
        platform_id: platform.platform_id,
        status: 'active',
        listed_at: listedAt,
        withdrawn_at: ''
      });
      logChange('Publications', id, 'CREATE', 'Recorded publication of item ' + item.reference_no + ' on ' + platform.name);
    });
    return { success: true, count: platforms.length };
  } finally { lock.releaseLock(); }
}

function withdrawPublication(publicationId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const publication = getRecords('Publications').find(row => String(row.publication_id) === String(publicationId));
    if (!publication) throw new Error('PUBLICATION_NOT_FOUND');
    if (!['active', 'remove_needed'].includes(String(publication.status).toLowerCase())) throw new Error('PUBLICATION_ALREADY_WITHDRAWN');
    const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!updateRecordFields('Publications', 'publication_id', publicationId, { status: 'withdrawn', withdrawn_at: date })) throw new Error('PUBLICATION_UPDATE_FAILED');
    logChange('Publications', publicationId, 'WITHDRAW', 'Withdrew publication');
    return { success: true };
  } finally { lock.releaseLock(); }
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

// Accepts a VAT percentage such as 4, 21, or a custom percentage.
function parseVatRate(value) {
  const normalized = String(value == null ? '' : value)
    .trim()
    .replace(',', '.');

  if (!/^\d+(\.\d{1,2})?$/.test(normalized)) {
    throw new Error('VAT_RATE_INVALID');
  }

  const rate = Number(normalized);
  if (!Number.isFinite(rate) || rate < 0 || rate > 100) {
    throw new Error('VAT_RATE_INVALID');
  }

  return rate;
}

function isValidSaleDate(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return false;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));

  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day;
}

function updateRecordFields(sheetName, keyField, keyValue, updates) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return false;

  const headers = values[0];
  const keyColumn = headers.indexOf(keyField);
  if (keyColumn < 0) return false;

  const rowIndex = values.findIndex((row, index) =>
    index > 0 && String(row[keyColumn]) === String(keyValue)
  );
  if (rowIndex < 1) return false;

  Object.entries(updates).forEach(([field, value]) => {
    const column = headers.indexOf(field);
    if (column < 0) {
      throw new Error('No s’ha trobat la columna ' + field + '.');
    }
    sheet.getRange(rowIndex + 1, column + 1).setValue(value);
  });

  return true;
}

function saveSale(form) {
  if (!form || !String(form.item_id || '').trim()) {
    throw new Error('SALE_ITEM_REQUIRED');
  }

  const saleDate = String(form.sale_date || '').trim();
  if (!isValidSaleDate(saleDate)) {
    throw new Error('SALE_DATE_REQUIRED');
  }

  const salesChannelType = String(form.sales_channel_type || '').trim();
  const salesChannelValue = String(form.sales_channel || '').trim();
  if (!salesChannelType || !salesChannelValue) throw new Error('SALE_CHANNEL_REQUIRED');

  let grossPrice;
  try {
    grossPrice = parsePrice(form.gross_price);
  } catch (error) {
    throw new Error('SALE_PRICE_INVALID');
  }
  if (grossPrice === '' || !Number.isFinite(grossPrice) || grossPrice < 0) {
    throw new Error('SALE_PRICE_INVALID');
  }

  const vatRate = parseVatRate(form.vat_rate);
  const vatAmount = Math.round(
    (grossPrice * vatRate / (100 + vatRate) + Number.EPSILON) * 100
  ) / 100;
  const netAmount = Math.round(
    (grossPrice - vatAmount + Number.EPSILON) * 100
  ) / 100;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const item = getRecords('Items').find(record =>
      String(record.item_id) === String(form.item_id)
    );

    if (!item) {
      throw new Error('ITEM_NOT_FOUND');
    }
    if (String(item.status).toLowerCase() !== 'active') {
      throw new Error('ITEM_NOT_ACTIVE');
    }

    const publications = getRecords('Publications');
    let salesChannel = salesChannelValue;
    if (salesChannelType === 'platform') {
      const platform = getRecords('Platforms').find(record =>
        String(record.platform_id) === salesChannelValue && String(record.active).toUpperCase() !== 'FALSE'
      );
      const matching = publications.filter(record =>
        String(record.item_id) === String(item.item_id) &&
        String(record.platform_id) === salesChannelValue &&
        String(record.status).toLowerCase() === 'active'
      );
      if (!platform || !matching.length) throw new Error('SALE_PLATFORM_INVALID');
      salesChannel = platform.name;
    } else if (salesChannelType !== 'other') {
      throw new Error('SALE_CHANNEL_REQUIRED');
    }
    const publicationsToRemove = publications.filter(record =>
      String(record.item_id) === String(item.item_id) && String(record.status).toLowerCase() === 'active'
    ).map(record => record.publication_id);

    const saleId = Utilities.getUuid();
    const sale = {
      sale_id: saleId,
      item_id: item.item_id,
      reference_no: item.reference_no,
      sale_date: saleDate,
      sales_channel: salesChannel,
      gross_price: grossPrice,
      vat_rate: vatRate,
      vat_amount: vatAmount,
      net_amount: netAmount,
      invoice_created: false,
      invoice_reference: '',
      status: 'active'
    };

    appendRecord('Sales', sale);

    try {
      publicationsToRemove.forEach(publicationId => {
        if (!updateRecordFields('Publications', 'publication_id', publicationId, { status: 'remove_needed' })) {
          throw new Error('PUBLICATION_UPDATE_FAILED');
        }
      });
      if (!updateRecordFields('Items', 'item_id', item.item_id, {
        status: 'sold',
        updated_at: new Date()
      })) {
        throw new Error('ITEM_UPDATE_FAILED');
      }
    } catch (error) {
      publicationsToRemove.forEach(publicationId => updateRecordFields('Publications', 'publication_id', publicationId, { status: 'active' }));
      updateRecordFields('Sales', 'sale_id', saleId, {
        status: 'reversed'
      });
      throw error;
    }

    logChange(
      'Sales',
      saleId,
      'CREATE',
      'Recorded sale for item ' + item.reference_no
    );

    return {
      success: true,
      reference_no: item.reference_no
    };
  } finally {
    lock.releaseLock();
  }
}

function reverseSale(saleId) {
  if (!saleId) {
    throw new Error('SALE_NOT_FOUND');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sale = getRecords('Sales').find(record =>
      String(record.sale_id) === String(saleId)
    );

    if (!sale) {
      throw new Error('SALE_NOT_FOUND');
    }
    if (String(sale.status).toLowerCase() !== 'active') {
      throw new Error('SALE_ALREADY_REVERSED');
    }

    const item = getRecords('Items').find(record =>
      String(record.item_id) === String(sale.item_id)
    );
    if (!item) {
      throw new Error('ITEM_NOT_FOUND');
    }
    if (String(item.status).toLowerCase() !== 'sold') {
      throw new Error('ITEM_NOT_SOLD');
    }

    if (!updateRecordFields('Items', 'item_id', item.item_id, {
      status: 'active',
      updated_at: new Date()
    })) {
      throw new Error('ITEM_UPDATE_FAILED');
    }

    try {
      if (!updateRecordFields('Sales', 'sale_id', sale.sale_id, {
        status: 'reversed'
      })) {
        throw new Error('SALE_UPDATE_FAILED');
      }
    } catch (error) {
      updateRecordFields('Items', 'item_id', item.item_id, {
        status: 'sold',
        updated_at: new Date()
      });
      throw error;
    }

    getRecords('Publications').filter(record =>
      String(record.item_id) === String(item.item_id) && String(record.status).toLowerCase() === 'remove_needed'
    ).forEach(record => updateRecordFields('Publications', 'publication_id', record.publication_id, { status: 'active' }));

    logChange(
      'Sales',
      sale.sale_id,
      'REVERSE',
      'Reversed sale for item ' + sale.reference_no
    );

    return {
      success: true,
      reference_no: sale.reference_no
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
