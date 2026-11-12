function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeIbsherCatalogKey(value) {
  return cleanText(value)
    .normalize("NFKC")
    .toLocaleLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function getLocalizedIbsherNames(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return {
      en: cleanText(value.en || value.english),
      ar: cleanText(value.ar || value.arabic),
      ku: cleanText(value.ku || value.kurdish),
    };
  }

  return { en: cleanText(value), ar: "", ku: "" };
}

function getIbsherApiList(payload, keys = []) {
  if (Array.isArray(payload)) return payload;
  for (const key of keys) {
    if (Array.isArray(payload?.[key])) return payload[key];
    if (Array.isArray(payload?.data?.[key])) return payload.data[key];
    const nested = payload?.[key];
    if (nested && typeof nested === "object") {
      for (const nestedKey of ["data", "results", "docs", key]) {
        if (Array.isArray(nested[nestedKey])) return nested[nestedKey];
      }
    }
  }
  return [];
}

function readIbsherListPagination(payload = {}) {
  const pagination = payload?.pagination || payload?.data?.pagination || {};
  return {
    currentPage: Number(pagination.currentPage || payload?.currentPage) || 0,
    totalPages: Number(pagination.totalPages || payload?.totalPages) || 0,
    total: Number(
      pagination.total
      || pagination.totalItems
      || pagination.totalProducts
      || payload?.total
      || payload?.totalItems
    ) || 0,
    limit: Number(pagination.limit || payload?.limit) || 0,
    hasPrevPage: pagination.hasPrevPage === true,
    hasNextPage: pagination.hasNextPage === true,
  };
}

function ibsherListHasNextPage(pagination = {}, { page = 1, collectedCount = 0 } = {}) {
  if (pagination.hasNextPage === true) return true;
  if (pagination.totalPages > 0 && page < pagination.totalPages) return true;
  if (pagination.total > 0 && collectedCount < pagination.total) return true;
  return false;
}

async function collectIbsherApiListPages({
  requestPage,
  keys = [],
  extractList,
  limit = 100,
  maxPages = 1000,
  getItemKey = (entry) => cleanText(entry?._id || entry?.id),
  incompleteMessage = "Could not load the complete IBSHER list.",
  tooLargeMessage = "The IBSHER list is too large to load completely.",
} = {}) {
  if (typeof requestPage !== "function") {
    throw new Error("requestPage is required to load IBSHER list pages.");
  }
  const readList = typeof extractList === "function"
    ? extractList
    : (payload) => getIbsherApiList(payload, keys);
  const items = [];
  const seen = new Set();

  for (let page = 1; page <= maxPages; page += 1) {
    const payload = await requestPage(page, limit);
    const entries = readList(payload);
    if (!Array.isArray(entries)) {
      throw new Error(incompleteMessage);
    }
    let freshCount = 0;
    for (const entry of entries) {
      const key = cleanText(getItemKey(entry, items.length)) || `row-${items.length}`;
      if (seen.has(key)) continue;
      seen.add(key);
      items.push(entry);
      freshCount += 1;
    }
    const pagination = readIbsherListPagination(payload);
    if (!ibsherListHasNextPage(pagination, { page, collectedCount: items.length })) {
      return items;
    }
    if (!freshCount) throw new Error(incompleteMessage);
  }

  throw new Error(tooLargeMessage);
}

function entityId(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return cleanText(value._id || value.id);
  }
  return cleanText(value);
}

function localizedEntity(entity, fallbackLabel) {
  const names = getLocalizedIbsherNames(entity?.name || entity?.label || fallbackLabel);
  const name = names.en || names.ar || names.ku;
  return {
    id: entityId(entity),
    name,
    aliases: [...new Set([names.en, names.ar, names.ku].map(cleanText).filter(Boolean))]
      .filter((value) => normalizeIbsherCatalogKey(value) !== normalizeIbsherCatalogKey(name)),
  };
}

function embeddedReference(reference, indexedEntities) {
  const id = entityId(reference);
  const indexed = id ? indexedEntities.get(id) : null;
  if (indexed) return indexed;
  if (!reference || typeof reference !== "object" || Array.isArray(reference)) {
    return { id, name: "", aliases: [] };
  }
  return localizedEntity(reference);
}

function buildIbsherProductCatalog(categoryPayload, subCategoryPayload, collectionPayload) {
  const rawCategories = getIbsherApiList(categoryPayload, ["category", "categories", "data"]);
  const rawSubCategories = getIbsherApiList(subCategoryPayload, [
    "subCategory",
    "subCategories",
    "subcategory",
    "subcategories",
    "data",
  ]);
  const rawCollections = getIbsherApiList(collectionPayload, ["collections", "collection", "data"]);

  const categoryById = new Map();
  const categories = [];
  for (const rawCategory of rawCategories) {
    if (!rawCategory || rawCategory.isActive === false) continue;
    const category = localizedEntity(rawCategory);
    if (!category.id || !category.name) continue;
    categoryById.set(category.id, category);
    categories.push(category);
  }

  const subCategoryById = new Map();
  const subCategories = [];
  for (const rawSubCategory of rawSubCategories) {
    if (!rawSubCategory || rawSubCategory.isActive === false) continue;
    const subCategory = localizedEntity(rawSubCategory);
    if (!subCategory.id || !subCategory.name) continue;
    const categoryReference = rawSubCategory.category ?? rawSubCategory.categoryId;
    const category = embeddedReference(categoryReference, categoryById);
    const record = {
      ...subCategory,
      categoryId: category.id,
      category: category.name,
    };
    subCategoryById.set(record.id, record);
    subCategories.push(record);
  }

  const collections = [];
  for (const rawCollection of rawCollections) {
    if (!rawCollection || rawCollection.isActive === false) continue;
    const collection = localizedEntity(rawCollection);
    if (!collection.id || !collection.name) continue;

    const subCategoryReference = rawCollection.subCategory ?? rawCollection.subCategoryId;
    const subCategory = embeddedReference(subCategoryReference, subCategoryById);
    const embeddedSubCategory = subCategoryReference && typeof subCategoryReference === "object"
      ? subCategoryReference
      : null;
    const categoryReference = embeddedSubCategory?.category
      ?? embeddedSubCategory?.categoryId
      ?? rawCollection.category
      ?? rawCollection.categoryId
      ?? subCategory.categoryId;
    const category = embeddedReference(categoryReference, categoryById);

    collections.push({
      ...collection,
      subCategoryId: subCategory.id,
      subCategory: subCategory.name,
      categoryId: category.id || subCategory.categoryId || "",
      category: category.name || subCategory.category || "",
    });
  }

  const byName = (left, right) => left.name.localeCompare(right.name);
  return {
    categories: categories.sort(byName),
    subCategories: subCategories.sort(byName),
    collections: collections.sort(byName),
  };
}

function sanitizeIbsherProductCatalog(catalog = {}) {
  const categories = Array.isArray(catalog.categories)
    ? catalog.categories.map((item) => (
      typeof item === "string" ? { id: "", name: cleanText(item), aliases: [] } : {
        id: entityId(item),
        name: cleanText(item?.name),
        aliases: Array.isArray(item?.aliases) ? item.aliases.map(cleanText).filter(Boolean) : [],
      }
    )).filter((item) => item.name)
    : [];
  const subCategories = Array.isArray(catalog.subCategories)
    ? catalog.subCategories.map((item) => ({
      id: entityId(item),
      name: cleanText(item?.name),
      aliases: Array.isArray(item?.aliases) ? item.aliases.map(cleanText).filter(Boolean) : [],
      categoryId: cleanText(item?.categoryId),
      category: cleanText(item?.category),
    })).filter((item) => item.name)
    : [];
  const collections = Array.isArray(catalog.collections)
    ? catalog.collections.map((item) => (
      typeof item === "string" ? {
        id: "",
        name: cleanText(item),
        aliases: [],
        categoryId: "",
        category: "",
        subCategoryId: "",
        subCategory: "",
      } : {
        id: entityId(item),
        name: cleanText(item?.name),
        aliases: Array.isArray(item?.aliases) ? item.aliases.map(cleanText).filter(Boolean) : [],
        categoryId: cleanText(item?.categoryId),
        category: cleanText(item?.category),
        subCategoryId: cleanText(item?.subCategoryId),
        subCategory: cleanText(item?.subCategory),
      }
    )).filter((item) => item.name)
    : [];

  return { categories, subCategories, collections };
}

function resolveIbsherCollection(value, catalog = {}) {
  const candidate = normalizeIbsherCatalogKey(value);
  if (!candidate) return null;
  const sanitized = sanitizeIbsherProductCatalog(catalog);
  const matches = sanitized.collections.filter((collection) => (
    [collection.name, ...collection.aliases]
      .some((name) => normalizeIbsherCatalogKey(name) === candidate)
  ));
  if (!matches.length) return null;

  const distinctIds = new Set(matches.map((item) => item.id).filter(Boolean));
  const distinctPaths = new Set(matches.map((item) => (
    `${item.categoryId || item.category}\0${item.subCategoryId || item.subCategory}\0${item.id || item.name}`
  )));
  return {
    ...matches[0],
    ambiguous: distinctIds.size > 1 || distinctPaths.size > 1,
  };
}

module.exports = {
  buildIbsherProductCatalog,
  collectIbsherApiListPages,
  entityId,
  getIbsherApiList,
  ibsherListHasNextPage,
  normalizeIbsherCatalogKey,
  readIbsherListPagination,
  resolveIbsherCollection,
  sanitizeIbsherProductCatalog,
};
