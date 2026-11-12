const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const xlsx = require('xlsx');
const ibsherDataExport = require('../public/lib/ibsher-data-export');

function product(brand) {
  return {
    _id: 'product-1', itemCode: 'SKU-1', brand, price: 20, discountPrice: 5,
    name: { en: 'Product' }, category: { _id: 'category-1' },
    variants: [{ color: { en: 'Black' }, sizes: [
      { size: 'S', stockStatus: 'in_stock', stockQuantity: 3 },
      { size: 'M', stockStatus: 'out_of_stock', stockQuantity: 0 },
    ], images: [{ url: '/photo.png' }] }],
  };
}

function roundTrip(exportData) {
  const workbook = ibsherDataExport.createIbsherExportWorkbook(xlsx, exportData);
  const bytes = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const read = xlsx.read(bytes, { type: 'buffer' });
  return xlsx.utils.sheet_to_json(read.Sheets.Sheet1, { header: 1, defval: '' });
}

test('Excel includes brand names for both string IDs and unpopulated brand objects', () => {
  const brands = [{ _id: 'brand-1', name: { en: 'Nike', ku: 'نایک' } }];
  for (const brand of ['brand-1', { _id: 'brand-1' }, { id: 'brand-1' }]) {
    const data = ibsherDataExport.buildIbsherExportRows([product(brand)], {}, brands);
    const [headers, ...rows] = roundTrip(data);
    assert.equal(rows.length, 2);
    assert.equal(headers.filter((header) => header === 'brand').length, 1);
    assert.deepEqual(rows.map((row) => row[headers.indexOf('brand')]), ['Nike', 'Nike']);
    assert.equal(rows[0][headers.indexOf('image-1')], '/photo.png');
  }
});

test('populated localized or plain brand names survive Excel export', () => {
  for (const [name, expected] of [
    ['Samsung', 'Samsung'], [{ en: 'Adidas' }, 'Adidas'],
    [{ en: ' ', ku: 'براندی کوردی' }, 'براندی کوردی'], [{ ar: 'علامة عربية' }, 'علامة عربية'],
  ]) {
    const data = ibsherDataExport.buildIbsherExportRows([product({ _id: 'brand-1', name })]);
    const [headers, row] = roundTrip(data);
    assert.equal(row[headers.indexOf('brand')], expected);
  }
});

test('inactive brands are resolved and brand, category and stock filters remain correct', () => {
  const brands = [{ id: 'brand-1', label: 'Legacy Brand', isActive: false }];
  const matching = product('brand-1');
  const other = { ...product('brand-2'), _id: 'product-2', itemCode: 'SKU-2' };
  const data = ibsherDataExport.buildIbsherExportRows([matching, other], {
    brandId: 'brand-1', categoryId: 'category-1', stock: 'in_stock',
  }, brands);
  assert.equal(data.productCount, 1);
  assert.equal(data.rows.length, 1);
  assert.equal(data.rows[0].brand, 'Legacy Brand');
  assert.equal(data.rows[0].size, 'S');
  assert.equal(data.rows[0].stockQuantity, 3);
});

test('missing or unknown brands remain blank without inventing a name or exporting an ID', () => {
  for (const brand of [null, undefined, 'unknown-id', { _id: 'unknown-id' }]) {
    const data = ibsherDataExport.buildIbsherExportRows([product(brand)]);
    assert.equal(data.rows[0].brand, '');
  }
});

test('the production export fetches brand names with the selected vendor and writes them to Excel', async () => {
  const source = fs.readFileSync(path.join(__dirname, '../public/electron.js'), 'utf8');
  const credentials = { value: 'vendor-1' };
  let brandPages = 0;
  let exported;
  const context = {
    ibsherDataExport,
    collectIbsherApiListPages: require('../public/lib/ibsher-catalog').collectIbsherApiListPages,
    cleanCell: (value) => String(value ?? '').trim(),
    requireAccountSectionAccess: (section) => assert.equal(section, 'ibsher-export'),
    resolveIbsherVendorAccount: async (options) => { assert.equal(options.vendorAccountValue, 'vendor-1'); return credentials; },
    getAllIbsherAdminProducts: async (options) => { assert.equal(options.credentials, credentials); return [product('brand-1')]; },
    requestIbsherAdmin: async (request, account) => {
      assert.equal(account, credentials);
      assert.equal(request.method, 'GET');
      assert.equal(request.url, '/brand');
      assert.equal(request.params.page, 1);
      assert.ok(request.params.limit >= 1);
      brandPages += 1;
      return {
        data: {
          brands: [{ _id: 'brand-1', name: { en: 'Nike' }, isActive: false }],
          pagination: { currentPage: 1, totalPages: 1, total: 1, hasNextPage: false },
        },
      };
    },
    dialog: { showSaveDialog: async () => ({ filePath: 'mock-export.xlsx' }) }, mainWindow: null,
    loadXlsx: () => ({ ...xlsx, writeFile: (workbook) => { exported = workbook; } }),
  };
  vm.createContext(context);
  const helperStart = source.indexOf('async function fetchIbsherAdminListPages(');
  const helperEnd = source.indexOf('function sanitizeIbsherProductCatalog(');
  assert.ok(helperStart >= 0 && helperEnd > helperStart);
  vm.runInContext(source.slice(helperStart, helperEnd), context);
  vm.runInContext(source.slice(source.indexOf('async function exportIbsherProducts('), source.indexOf('async function selectIbsherDuplicateExcelFile(')), context);
  const result = await context.exportIbsherProducts({ vendorAccountValue: 'vendor-1' });
  assert.equal(result.exportedProducts, 1);
  assert.equal(brandPages, 1);
  const rows = xlsx.utils.sheet_to_json(exported.Sheets.Sheet1);
  assert.deepEqual(rows.map((row) => row.brand), ['Nike', 'Nike']);
});
