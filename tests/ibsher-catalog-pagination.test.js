const test = require('node:test');
const assert = require('node:assert/strict');
const {
  buildIbsherProductCatalog,
  collectIbsherApiListPages,
  getIbsherApiList,
  ibsherListHasNextPage,
  readIbsherListPagination,
} = require('../public/lib/ibsher-catalog');

test('getIbsherApiList reads nested data.brands style payloads', () => {
  assert.deepEqual(
    getIbsherApiList({ data: { brands: [{ _id: 'b1' }] } }, ['brands', 'brand', 'data']),
    [{ _id: 'b1' }],
  );
});

test('readIbsherListPagination matches the brand-import pagination contract', () => {
  assert.deepEqual(readIbsherListPagination({
    brands: [],
    pagination: { currentPage: 2, totalPages: 4, total: 350, limit: 100, hasNextPage: true },
  }), {
    currentPage: 2,
    totalPages: 4,
    total: 350,
    limit: 100,
    hasPrevPage: false,
    hasNextPage: true,
  });
  assert.equal(
    ibsherListHasNextPage(readIbsherListPagination({ pagination: { totalPages: 3 } }), { page: 2, collectedCount: 200 }),
    true,
  );
  assert.equal(
    ibsherListHasNextPage(readIbsherListPagination({ pagination: { total: 5 } }), { page: 1, collectedCount: 5 }),
    false,
  );
});

test('collectIbsherApiListPages merges pages until hasNextPage clears', async () => {
  const seen = [];
  const items = await collectIbsherApiListPages({
    keys: ['subCategory', 'data'],
    limit: 2,
    requestPage: async (page, limit) => {
      seen.push({ page, limit });
      if (page === 1) {
        return {
          subCategory: [{ _id: 's1', name: { en: 'One' } }, { _id: 's2', name: { en: 'Two' } }],
          pagination: { currentPage: 1, totalPages: 2, total: 3, hasNextPage: true },
        };
      }
      return {
        subCategory: [{ _id: 's3', name: { en: 'Three' } }],
        pagination: { currentPage: 2, totalPages: 2, total: 3, hasNextPage: false },
      };
    },
  });
  assert.deepEqual(seen, [{ page: 1, limit: 2 }, { page: 2, limit: 2 }]);
  assert.deepEqual(items.map((item) => item._id), ['s1', 's2', 's3']);
});

test('buildIbsherProductCatalog accepts already-merged subcategory and collection arrays', () => {
  const catalog = buildIbsherProductCatalog(
    { category: [{ _id: 'c1', name: { en: 'Clothes' }, isActive: true }] },
    { subCategory: [{ _id: 'sc1', name: { en: 'Tops' }, category: 'c1', isActive: true }] },
    { collections: [{ _id: 'col1', name: { en: 'Summer' }, subCategory: 'sc1', isActive: true }] },
  );
  assert.equal(catalog.categories.length, 1);
  assert.equal(catalog.subCategories[0].categoryId, 'c1');
  assert.equal(catalog.collections[0].subCategoryId, 'sc1');
});
