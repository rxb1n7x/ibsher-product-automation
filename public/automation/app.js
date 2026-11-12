import puppeteer from "puppeteer-core";
import xlsx from "xlsx";
import path from "path";
import fs from "fs";
import axios from "axios";
import FormData from "form-data";
import readline from "readline";
import { Writable } from "stream";
import variantBuilder from "./product-variants.cjs";
import productRowSorter from "./product-row-sort.cjs";
import localImageSource from "./local-image-source.cjs";
import apiImageFormat from "./api-image-format.cjs";
import ibsherCatalog from "../lib/ibsher-catalog.js";
import automationReportExporter from "../lib/automation-report-export.js";

const {
  buildApiVariants,
  buildVariantImageFieldName,
  createProductFromRows,
  findProductRowConflicts,
  flattenProductRows,
  normalizeProductPricing,
} = variantBuilder;
const { normalizeItemcodeField, sortProductRowsByItemcode } = productRowSorter;
const { resolveLocalImageSource } = localImageSource;
const { prepareApiUploadImage } = apiImageFormat;
const {
  buildIbsherProductCatalog,
  collectIbsherApiListPages,
  getIbsherApiList,
  resolveIbsherCollection,
} = ibsherCatalog;
const { buildAutomationReportItems } = automationReportExporter;


// ===== CONFIG =====
const EXCEL_FILE = process.env.EXCEL_FILE || "./data-entry.xlsx";
const SOURCE_EXCEL_FILE = process.env.SOURCE_EXCEL_FILE || EXCEL_FILE;
const IMAGE_SOURCE_DIR = cleanText(process.env.IMAGE_SOURCE_DIR);
const SHEET_NAME = process.env.SHEET_NAME || "Sheet1";
const DOWNLOAD_DIR = process.env.DOWNLOAD_DIR || path.join("./temp_images");
const AUTOMATION_REPORT_PREFIX = "__IBSHER_AUTOMATION_REPORT__";
const MAX_IMAGES_PER_PRODUCT = Math.max(1, Number.parseInt(process.env.MAX_IMAGES_PER_PRODUCT || "15", 10) || 15);
const AUTOMATION_MODE = cleanText(process.env.AUTOMATION_MODE).toLowerCase() === "browser" ? "browser" : "api";
const API_BASE_URL = cleanText(process.env.IBSHER_API_BASE_URL) || "https://api.ibsher.com/api/v1/admin";
const API_CONCURRENCY = Math.min(5, Math.max(1, Number.parseInt(process.env.API_CONCURRENCY || "3", 10) || 3));
const IS_INTERACTIVE_TERMINAL = Boolean(process.stdin.isTTY && process.stdout.isTTY);
const IBSHER_ITEM_CODE_FIELD = "itemCode";

// Create temp directory if it doesn't exist
if (!fs.existsSync(DOWNLOAD_DIR)) {
  fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
}

// ===== READ EXCEL =====
const workbook = xlsx.readFile(EXCEL_FILE);
const sheet = workbook.Sheets[SHEET_NAME];

if (!sheet) {
  throw new Error(`Sheet "${SHEET_NAME}" not found in ${EXCEL_FILE}`);
}

const rows = sortProductRowsByItemcode(
  xlsx.utils.sheet_to_json(sheet).map((row) => normalizeProductPricing(normalizeItemcodeField(row)))
);

// ===== TRACK FAILED ITEMS =====
const failedItems = {
  missingDetails: [],
  noBrand: [],
  noCollection: [],
  noBrandAndCollection: [],
  dropdownNotFound: [],
  duplicateItemcode: [],
  invalidProductInfo: [],
  imageDownloadFailures: [],
  otherErrors: []
};

// ===== PRODUCT VALIDATION =====
const REQUIRED_PRODUCT_FIELDS = [
  ["itemcode", "item code"],
  ["en", "English name"],
  ["ar", "Arabic name"],
  ["ku", "Kurdish name"],
  ["collectionName", "collection"],
  ["brand", "brand"],
  ["description-en", "English description"],
  ["description-ar", "Arabic description"],
  ["description-ku", "Kurdish description"],
  ["price", "price"],
  ["stockQuantity", "stock quantity"],
];

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeStockStatus(value) {
  const cleaned = cleanText(value);
  const comparable = cleaned.toLowerCase().replace(/[\s_-]+/g, " ").trim();
  return comparable === "out of stock" ? "out_of_stock" : cleaned;
}

function isOutOfStockQuantity(value) {
  return normalizeStockStatus(value) === "out_of_stock";
}

function isMissing(value) {
  return value === undefined || value === null || cleanText(value) === "";
}

function getExcelRowNumber(row, rowIndex) {
  return Number.isInteger(row.__rowNum__) ? row.__rowNum__ + 1 : rowIndex + 2;
}

function parseNumber(value) {
  if (typeof value === "number") return value;
  return Number(cleanText(value).replace(/,/g, ""));
}

function createSkippedItem(row, rowIndex, reason, extra = {}) {
  return {
    rowNumber: getExcelRowNumber(row, rowIndex),
    itemcode: cleanText(row.itemcode) || "(missing)",
    name: cleanText(row.en) || "(missing)",
    reason,
    ...extra,
  };
}

function getMissingRequiredFields(row) {
  return REQUIRED_PRODUCT_FIELDS
    .filter(([field]) => isMissing(row[field]))
    .map(([, label]) => label);
}

function getImageUrls(row) {
  const images = [];
  for (let i = 1; i <= 15; i++) {
    const imgPath = cleanText(row[`image-${i}`]);
    if (imgPath) {
      images.push(imgPath);
    }
  }
  return images;
}

function getInvalidProductInfo(row) {
  const errors = [];
  const price = parseNumber(row.price);
  const stockQuantity = parseNumber(row.stockQuantity);
  const isOutOfStock = isOutOfStockQuantity(row.stockQuantity);

  if (row.__priceCurrencyConflict) {
    errors.push("enter a value in only one of price USD or price IQD");
  }

  if (!Number.isFinite(price) || price <= 0) {
    errors.push("price must be a number greater than 0");
  }

  if (!isOutOfStock && (!Number.isInteger(stockQuantity) || stockQuantity < 0)) {
    errors.push("stock quantity must be a whole number of 0 or more, or Out Of Stock");
  }

  if (!isMissing(row.discountPrice)) {
    const discountPrice = parseNumber(row.discountPrice);
    if (!Number.isFinite(discountPrice) || discountPrice < 0) {
      errors.push("discount price must be a number of 0 or more");
    } else if (Number.isFinite(price) && discountPrice >= price) {
      errors.push("discount price must be less than price");
    }
  }

  if (!isMissing(row.warrantyDuration)) {
    const warrantyDuration = parseNumber(row.warrantyDuration);
    if (!Number.isFinite(warrantyDuration) || warrantyDuration <= 0) {
      errors.push("warranty duration must be a number greater than 0");
    }
    if (isMissing(row.warrantyUnit)) {
      errors.push("warranty unit is required when warranty duration is provided");
    }
  }

  return errors;
}

// ===== BUILD PRODUCTS, GROUPING DUPLICATE ITEM CODES AS COLOR OPTIONS =====
const itemcodeCounts = rows.reduce((acc, row) => {
  const itemcodeKey = cleanText(row.itemcode).toLocaleLowerCase();
  if (itemcodeKey) {
    acc.set(itemcodeKey, (acc.get(itemcodeKey) || 0) + 1);
  }
  return acc;
}, new Map());

const validRowsByItemcode = new Map();
const conflictsByItemcode = findProductRowConflicts(rows);
let validRowCount = 0;

for (const [rowIndex, row] of rows.entries()) {
  const itemcode = cleanText(row.itemcode);
  const itemcodeKey = itemcode.toLocaleLowerCase();
  const missingFields = getMissingRequiredFields(row);
  if (itemcodeCounts.get(itemcodeKey) > 1) {
    if (isMissing(row.color)) missingFields.push("color");
    if (isMissing(row.size)) missingFields.push("size");
    if (!getImageUrls(row).length) missingFields.push("image for this color option");
  }

  if (missingFields.length > 0) {
    const missingBrand = missingFields.includes("brand");
    const missingCollection = missingFields.includes("collection");
    const skippedItem = createSkippedItem(
      row,
      rowIndex,
      `Missing required details: ${missingFields.join(", ")}`,
      {
        brand: cleanText(row.brand),
        collection: cleanText(row.collectionName),
        missingFields,
      },
    );

    if (missingBrand && missingCollection) {
      failedItems.noBrandAndCollection.push(skippedItem);
    } else if (missingBrand) {
      failedItems.noBrand.push(skippedItem);
    } else if (missingCollection) {
      failedItems.noCollection.push(skippedItem);
    } else {
      failedItems.missingDetails.push(skippedItem);
    }

    continue;
  }

  const invalidInfo = getInvalidProductInfo(row);
  const conflictingColumns = conflictsByItemcode.get(itemcodeKey) || [];
  if (conflictingColumns.length) {
    invalidInfo.push(`rows with the same itemcode must match in: ${conflictingColumns.join(", ")}`);
  }
  if (invalidInfo.length > 0) {
    failedItems.invalidProductInfo.push(
      createSkippedItem(row, rowIndex, `Invalid product info: ${invalidInfo.join("; ")}`, {
        errors: invalidInfo,
      })
    );
    continue;
  }

  if (!validRowsByItemcode.has(itemcodeKey)) {
    validRowsByItemcode.set(itemcodeKey, { itemcode, rows: [] });
  }
  validRowsByItemcode.get(itemcodeKey).rows.push(row);
  validRowCount += 1;
}

const products = [...validRowsByItemcode.values()].map(({ itemcode, rows: productRows }) => (
  createProductFromRows(productRows, itemcode, getImageUrls)
));

const colorOptionCount = products.reduce((total, product) => total + product.variants.length, 0);
console.log(
  `\nGrouped ${validRowCount} valid row(s) into ${products.length} product(s) and ${colorOptionCount} color option(s). `
  + `Skipped ${rows.length - validRowCount} invalid row(s).`,
);

// ===== TRACK PROCESSED ITEM CODES =====
const processedItemcodes = new Set();

// ===== GET USER OPTION BEFORE STARTING =====
let submissionMode = 'manual'; // Default to manual

function getUserChoice() {
  const envSubmissionMode = cleanText(process.env.SUBMISSION_MODE).toLowerCase();
  if (envSubmissionMode === "auto" || envSubmissionMode === "manual") {
    return Promise.resolve(envSubmissionMode);
  }

  if (!IS_INTERACTIVE_TERMINAL) {
    return Promise.resolve("manual");
  }

  return new Promise((resolve) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });

    console.log("\n" + "=".repeat(60));
    console.log("🚀 PRODUCT SUBMISSION MODE");
    console.log("=".repeat(60));
    console.log("1. Auto Submit - Automatically submit each product form");
    console.log("2. Manual Submit - You will manually click Submit for each product");
    console.log("=".repeat(60));
    
    rl.question("\nSelect mode (1 or 2): ", (answer) => {
      rl.close();
      if (answer.trim() === '1') {
        resolve('auto');
      } else {
        resolve('manual');
      }
    });
  });
}

function askRequiredQuestion(prompt, options = {}) {
  const { secret = false } = options;

  return new Promise((resolve) => {
    const output = secret
      ? new Writable({
          write(chunk, encoding, callback) {
            callback();
          }
        })
      : process.stdout;

    const rl = readline.createInterface({
      input: process.stdin,
      output,
      terminal: true
    });

    if (secret) {
      process.stdout.write(prompt);
    }

    rl.question(secret ? "" : prompt, (answer) => {
      if (secret) {
        process.stdout.write("\n");
      }
      rl.close();

      const value = answer.trim();
      if (value) {
        resolve(value);
      } else {
        console.log("This field is required.");
        resolve(askRequiredQuestion(prompt, options));
      }
    });
  });
}

async function getIbsherCredentials() {
  const envEmail = cleanText(process.env.IBSHER_EMAIL);
  const envPassword = cleanText(process.env.IBSHER_PASSWORD);

  if (envEmail && envPassword) {
    console.log("\nIBSHER LOGIN");
    console.log("Using credentials provided by the Electron app.");
    return {
      email: envEmail,
      password: envPassword,
    };
  }

  if (!IS_INTERACTIVE_TERMINAL) {
    throw new Error("IBSHER_EMAIL and IBSHER_PASSWORD are required when running from Electron.");
  }

  console.log("\n" + "=".repeat(60));
  console.log("IBSHER LOGIN");
  console.log("=".repeat(60));

  return {
    email: await askRequiredQuestion("IBSHER_EMAIL: "),
    password: await askRequiredQuestion("IBSHER_PASSWORD: ", { secret: true })
  };
}
// ===== START BROWSER =====
const browserLaunchOptions = {
  headless: false,
  defaultViewport: null,
  args: ['--disable-web-security', '--disable-features=IsolateOrigins,site-per-process']
};

if (process.env.PUPPETEER_EXECUTABLE_PATH) {
  browserLaunchOptions.executablePath = process.env.PUPPETEER_EXECUTABLE_PATH;
} else {
  browserLaunchOptions.channel = process.env.PUPPETEER_CHANNEL || "chrome";
}

let browser = null;
let page = null;

if (AUTOMATION_MODE === "browser") {
  browser = await puppeteer.launch(browserLaunchOptions);
  page = await browser.newPage();
}

// ===== SELECTORS =====
const itemCodeInputSelector = `input[name="${IBSHER_ITEM_CODE_FIELD}"]`;
const addProductSelectors = [
  '[data-testid="add-product"]',
  'button[aria-label*="add product" i]',
  '[role="button"][aria-label*="add product" i]',
  // Legacy iBsher UI fallback. Keep this last because presentation classes are unstable.
  '.cursor-pointer.bg-gray-100.text-black.px-4.py-2.rounded-lg',
];

async function waitForAddProductForm() {
  // The current iBsher /add-product route renders the form directly. Older
  // versions first showed an Add Product button, so support both layouts.
  try {
    return await page.waitForSelector(itemCodeInputSelector, {
      visible: true,
      timeout: 10000,
    });
  } catch {
    // Fall through to the legacy/alternate trigger lookup.
  }

  let trigger = null;
  for (const selector of addProductSelectors) {
    trigger = await page.$(selector);
    if (trigger) break;
  }

  if (!trigger) {
    const handle = await page.evaluateHandle(() => {
      const candidates = [...document.querySelectorAll('button, [role="button"]')];
      return candidates.find((element) => {
        const rect = element.getBoundingClientRect();
        const style = window.getComputedStyle(element);
        const text = (element.innerText || element.textContent || element.getAttribute("aria-label") || "")
          .trim()
          .toLowerCase();
        return rect.width > 0
          && rect.height > 0
          && style.display !== "none"
          && style.visibility !== "hidden"
          && ["add product", "add products", "create product"].includes(text);
      }) || null;
    });
    trigger = handle.asElement();
  }

  if (!trigger) {
    throw new Error(
      `Add Product form did not load and no button to open it was found (URL: ${page.url()})`,
    );
  }

  await trigger.evaluate((element) => element.scrollIntoView({ block: "center" }));
  await trigger.click();
  return page.waitForSelector(itemCodeInputSelector, { visible: true, timeout: 10000 });
}

// ===== AUTO CANCEL FORM FUNCTION =====
async function autoCancelForm() {
  console.log("\n🔄 Auto-cancelling form...");
  
  try {
    // Try to find and click cancel button
    const cancelSelectors = [
      'button:has-text("Cancel")',
      'button:has-text("cancel")',
      'button:has-text("Close")',
      'button:has-text("close")',
      'button:has-text("Discard")',
      'button:has-text("discard")',
      'button[class*="cancel"]',
      'button[class*="close"]',
      'button[class*="discard"]',
      '.btn-cancel',
      '.btn-close',
      'button:has-text("✕")',
      'button:has-text("×")',
      '[aria-label="Close"]',
      '[aria-label="Cancel"]'
    ];
    
    let cancelButton = null;
    for (const selector of cancelSelectors) {
      try {
        cancelButton = await page.$(selector);
        if (cancelButton) {
          break;
        }
      } catch (e) {
        continue;
      }
    }
    
    if (cancelButton) {
      await cancelButton.click();
      console.log("✅ Cancel button clicked!");
      await new Promise(r => setTimeout(r, 1000));
      
      // Check if we need to confirm cancellation
      const confirmSelectors = [
        'button:has-text("Yes")',
        'button:has-text("Confirm")',
        'button:has-text("OK")',
        'button:has-text("ok")',
        'button[class*="confirm"]',
        '.btn-confirm',
        '.btn-primary:has-text("Yes")'
      ];
      
      let confirmButton = null;
      for (const selector of confirmSelectors) {
        try {
          confirmButton = await page.$(selector);
          if (confirmButton) {
            break;
          }
        } catch (e) {
          continue;
        }
      }
      
      if (confirmButton) {
        await confirmButton.click();
        console.log("✅ Confirmed cancellation!");
        await new Promise(r => setTimeout(r, 1000));
      }
      
      return true;
    } else {
      // If no cancel button, navigate back directly
      console.log("⚠️ No cancel button found, navigating back to add-product page...");
      await page.goto("https://marketplace.ibsher.com/add-product", {
        waitUntil: "networkidle2",
        timeout: 30000
      });
      return true;
    }
  } catch (error) {
    console.log(`❌ Auto-cancel error: ${error.message}`);
    // Fallback: navigate back directly
    try {
      await page.goto("https://marketplace.ibsher.com/add-product", {
        waitUntil: "networkidle2",
        timeout: 30000
      });
      return true;
    } catch (navError) {
      console.log(`❌ Failed to navigate back: ${navError.message}`);
      return false;
    }
  }
}

// ===== HELPER FUNCTIONS =====
async function waitForManualSubmit() {
  console.log("\n⏳ Waiting for you to manually click the SUBMIT button on the website...");
  
  return new Promise(async (resolve) => {
    let checkInterval = setInterval(async () => {
      try {
        const formExists = await page.$(`input[name="${IBSHER_ITEM_CODE_FIELD}"]`);
        const submitButtonExists = await page.$('button[type="submit"]');
        
        if (!formExists || !submitButtonExists) {
          clearInterval(checkInterval);
          console.log("✅ Form submission detected! Moving to next product...");
          resolve(true);
        }
      } catch (error) {
        clearInterval(checkInterval);
        console.log("✅ Form submission detected! Moving to next product...");
        resolve(true);
      }
    }, 1000);
    
    let currentUrl = page.url();
    const urlCheckInterval = setInterval(async () => {
      const newUrl = page.url();
      if (newUrl !== currentUrl) {
        clearInterval(urlCheckInterval);
        clearInterval(checkInterval);
        console.log("✅ Page navigated - form submitted! Moving to next product...");
        resolve(true);
      }
    }, 1000);
  });
}

async function autoSubmitForm() {
  try {
    console.log("\n🤖 Auto-submitting form...");
    
    // Try to find and click the submit button
    const submitSelectors = [
      'button[type="submit"]',
      'button:has-text("Submit")',
      'button:has-text("submit")',
      'button:has-text("Save")',
      'button:has-text("save")',
      'button:has-text("Add Product")',
      'button:has-text("add product")',
      '.btn-primary',
      '.btn-submit',
      'button[class*="submit"]',
      'button[class*="save"]'
    ];
    
    let submitButton = null;
    for (const selector of submitSelectors) {
      try {
        submitButton = await page.$(selector);
        if (submitButton) {
          break;
        }
      } catch (e) {
        continue;
      }
    }
    
    if (submitButton) {
      await submitButton.click();
      console.log("✅ Submit button clicked!");
      
      // Wait for navigation or form submission
      await new Promise(r => setTimeout(r, 3000));
      
      // Check if form was submitted successfully
      const currentUrl = page.url();
      if (currentUrl.includes('/add-product')) {
        return false;
      } else {
        console.log("✅ Form submitted successfully!");
        return true;
      }
    } else {
      console.log("❌ Could not find submit button!");
      return false;
    }
  } catch (error) {
    console.error(`❌ Auto-submit error: ${error.message}`);
    return false;
  }
}

async function clearAndType(selector, value) {
  try {
    const el = await page.waitForSelector(selector, { visible: true, timeout: 5000 });
    await el.click({ clickCount: 3 });
    await page.keyboard.press("Backspace");
    if (value !== undefined && value !== null && String(value) !== "") {
      await el.type(String(value));
    }
  } catch (error) {
    console.log(`Could not fill ${selector}:`, error.message);
  }
}

async function tryType(selector, value) {
  try {
    const el = await page.$(selector);
    if (!el) return false;
    await el.click({ clickCount: 3 });
    await page.keyboard.press("Backspace");
    if (value !== undefined && value !== null && String(value) !== "") {
      await el.type(String(value));
    }
    return true;
  } catch (error) {
    return false;
  }
}

// ===== REACT SELECT DROPDOWN FUNCTION =====
async function selectFromDropdown(inputHandle, searchValue, fieldName, options = {}) {
  const {
    exactMatch = false,
    waitForDropdown = true,
    dropdownTimeout = 3000,
    maxRetries = 2,
    optionValue = ""
  } = options;

  if (!searchValue || searchValue === "") {
    console.log(`⚠️ No ${fieldName} value provided, skipping...`);
    return { success: false, reason: "no_value" };
  }

  const searchLower = searchValue.toLowerCase();
  const optionValueLower = optionValue.toLowerCase();

  for (let retry = 0; retry <= maxRetries; retry++) {
    try {
      // Get the input element
      const inputElement = await inputHandle.asElement();
      if (!inputElement) {
        throw new Error("Input element not found");
      }

      // Click to focus and open dropdown
      await inputElement.click({ clickCount: 3 });
      await page.keyboard.press("Backspace");
      
      // Clear existing value
      await inputElement.click();
      await page.keyboard.down('Control');
      await page.keyboard.press('a');
      await page.keyboard.up('Control');
      await page.keyboard.press('Backspace');
      
      // Type the search value
      await inputElement.type(searchValue);
      
      // Wait for dropdown to appear
      await new Promise(r => setTimeout(r, 500));

      if (waitForDropdown) {
        try {
          // React Select specific selectors
          await page.waitForFunction(
            () => {
              const options = document.querySelectorAll('[role="option"], .css-1n7v3ny-option, .css-1s2u09g-option, [class*="option"]');
              return options.length > 0;
            },
            { timeout: dropdownTimeout }
          );
        } catch (waitError) {
          console.log(`⚠️ No dropdown appeared for ${fieldName}: "${searchValue}"`);
          // Return failure for dropdown not appearing
          return { success: false, reason: "dropdown_not_appeared", searchValue: searchValue };
        }
      }

      // React Select specific option selectors
      const optionSelectors = [
        '[role="option"]',
        '.css-1n7v3ny-option',  // React Select selected option
        '.css-1s2u09g-option',  // React Select option
        '.css-1h46j35-option',
        '[class*="option"]',
        '[class*="Option"]',
        '[data-value]',  // React Select data attribute
        'div[class*="menu"] div[class*="option"]'
      ];

      let selectedOption = null;
      let selectedText = null;

      // Find all options
      for (const selector of optionSelectors) {
        const options = await page.$$(selector);
        
        for (const option of options) {
          try {
            const isVisible = await option.evaluate(el => {
              const rect = el.getBoundingClientRect();
              return rect.width > 0 && rect.height > 0 && 
                     window.getComputedStyle(el).display !== 'none' &&
                     window.getComputedStyle(el).visibility !== 'hidden';
            });
            
            if (!isVisible) continue;

            const optionDetails = await option.evaluate(el => ({
              text: el.innerText?.trim() || el.textContent?.trim() || "",
              value: el.getAttribute('data-value') || el.getAttribute('value') || ""
            }));
            const text = optionDetails.text;
            const textLower = text.toLowerCase();
            const valueLower = optionDetails.value.toLowerCase();

            let isMatch = false;
            if (optionValueLower && valueLower === optionValueLower) {
              isMatch = true;
            } else if (exactMatch) {
              isMatch = textLower === searchLower;
            } else {
              isMatch = textLower.includes(searchLower) || searchLower.includes(textLower);
            }

            if (isMatch) {
              if (textLower === searchLower || !selectedOption) {
                selectedOption = option;
                selectedText = text;
                if (textLower === searchLower) break;
              }
            }
          } catch (e) {
            continue;
          }
        }
        
        if (selectedOption) break;
      }

      if (selectedOption) {
        // Scroll into view and click
        await selectedOption.evaluate(el => el.scrollIntoView({ block: 'nearest' }));
        await new Promise(r => setTimeout(r, 200));
        await selectedOption.click();
        console.log(`✅ Selected ${fieldName}: "${selectedText}"`);
        return { success: true, selectedText: selectedText };
      } else {
        if (exactMatch) {
          console.log(`⚠️ No exact option found for ${fieldName}: "${searchValue}"`);
          return { success: false, reason: "option_not_found", searchValue };
        }
        // Try keyboard navigation as fallback
        console.log(`⚠️ No matching option found for ${fieldName}: "${searchValue}", trying keyboard...`);
        
        // Press Enter to confirm the typed value (React Select may accept it)
        await page.keyboard.press("Enter");
        await new Promise(r => setTimeout(r, 300));
        
        console.log(`✅ ${fieldName} set via keyboard: ${searchValue}`);
        return { success: true, selectedText: searchValue };
      }

    } catch (error) {
      if (retry < maxRetries) {
        console.log(`⚠️ Retry ${retry + 1} for ${fieldName}...`);
        await new Promise(r => setTimeout(r, 500));
        continue;
      }
      console.error(`Error selecting ${fieldName}: ${error.message}`);
      return { success: false, reason: "error", error: error.message };
    }
  }

  return { success: false, reason: "max_retries" };
}

// ===== FIND INPUT FIELD HELPER =====
async function findInputField(fieldName, searchTerms) {
  // Try direct selector first
  for (const term of searchTerms) {
    const directSelector = `input[name="${term}"]`;
    let input = await page.$(directSelector);
    if (input) return input;
  }

  // Try evaluateHandle with multiple strategies
  const inputHandle = await page.evaluateHandle((searchTerms) => {
    // Strategy 1: Match by name attribute
    const allInputs = document.querySelectorAll('input');
    for (const input of allInputs) {
      const name = (input.name || '').toLowerCase();
      const placeholder = (input.placeholder || '').toLowerCase();
      
      for (const term of searchTerms) {
        const termLower = term.toLowerCase();
        if (name === termLower || name.includes(termLower) || placeholder.includes(termLower)) {
          return input;
        }
      }
    }

    // Strategy 2: Match by associated label
    const labels = document.querySelectorAll('label');
    for (const label of labels) {
      const labelText = (label.innerText || '').toLowerCase();
      for (const term of searchTerms) {
        if (labelText.includes(term.toLowerCase())) {
          const input = label.closest('div')?.querySelector('input') || 
                       label.parentElement?.querySelector('input') ||
                       document.getElementById(label.getAttribute('for'));
          if (input) return input;
        }
      }
    }

    // Strategy 3: Match by parent container text
    const containers = document.querySelectorAll('[class*="form"], [class*="field"], [class*="group"], .mb-4, .mb-6');
    for (const container of containers) {
      const containerText = (container.innerText || '').toLowerCase();
      for (const term of searchTerms) {
        if (containerText.includes(term.toLowerCase())) {
          const input = container.querySelector('input:not([type="file"]):not([type="checkbox"]):not([type="radio"]):not([type="hidden"])');
          if (input) return input;
        }
      }
    }

    return null;
  }, searchTerms);

  return inputHandle;
}

// ===== CATEGORY HIERARCHY FUNCTIONS (FOR REACT SELECT) =====
async function fillCatalogDropdown(value, fieldName, searchTerms) {
  if (!cleanText(value)) {
    console.log(`⚠️ No ${fieldName} provided`);
    return { success: false, reason: "no_value" };
  }

  try {
    const input = await findInputField(fieldName, searchTerms);
    const inputElement = input?.asElement ? input.asElement() : null;
    if (!inputElement) {
      console.log(`❌ Could not find ${fieldName} input field`);
      return { success: false, reason: "input_not_found" };
    }
    return selectFromDropdown(input, value, fieldName, {
      exactMatch: true,
      waitForDropdown: true,
      dropdownTimeout: 3000,
      maxRetries: 3,
    });
  } catch (error) {
    console.error(`Error filling ${fieldName}: ${error.message}`);
    return { success: false, reason: "error", error: error.message };
  }
}

function fillCategory(categoryName) {
  return fillCatalogDropdown(categoryName, "Category", ["category"]);
}

function fillSubCategory(subCategoryName) {
  return fillCatalogDropdown(subCategoryName, "Sub-category", [
    "subCategory",
    "subcategory",
    "sub-category",
    "sub category",
  ]);
}

function fillCollection(collectionName) {
  return fillCatalogDropdown(collectionName, "Collection", [
    "collectionName",
    "collection",
    "product type",
  ]);
}

async function fillCategoryHierarchy(collection) {
  const selections = [
    { field: "category", value: collection.category, select: fillCategory },
    { field: "subCategory", value: collection.subCategory, select: fillSubCategory },
    { field: "collection", value: collection.name, select: fillCollection },
  ];

  for (const selection of selections) {
    console.log(`\nSetting ${selection.field} with auto-select...`);
    const result = await selection.select(selection.value);
    if (!result.success) return { ...result, ...selection };
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  return { success: true };
}

// ===== BRAND FUNCTION (FOR REACT SELECT) =====
async function fillBrand(brandName) {
  if (!brandName || brandName === "") {
    console.log("⚠️ No brand name provided");
    return { success: false, reason: "no_brand" };
  }
  
  try {
    const brandInput = await findInputField("Brand", ["brand"]);
    const brandElement = brandInput?.asElement ? brandInput.asElement() : null;

    if (brandElement) {
      const result = await selectFromDropdown(brandInput, brandName, "Brand", {
        exactMatch: false,
        waitForDropdown: true,
        dropdownTimeout: 3000,
        maxRetries: 3
      });
      return result;
    } else {
      console.log(`❌ Could not find brand input field`);
      return { success: false, reason: "input_not_found" };
    }
  } catch (error) {
    console.error(`Error filling brand: ${error.message}`);
    return { success: false, reason: "error", error: error.message };
  }
}

// ===== STOCK STATUS FUNCTION =====
async function fillStockStatus(statusValue, variantIndex = 0, sizeIndex = 0) {
  const statusCode = normalizeStockStatus(statusValue);
  if (!statusCode) {
    return { success: true, skipped: true };
  }

  // The importer currently supports the requested out-of-stock state only.
  if (statusCode !== "out_of_stock") {
    console.log(`Skipping unsupported stock status: "${statusValue}"`);
    return { success: true, skipped: true };
  }

  try {
    const nativeSelectResult = await page.evaluate(({ value, variantIndex, sizeIndex }) => {
      const comparable = (text) => String(text || "")
        .toLowerCase()
        .replace(/[\s_-]+/g, " ")
        .trim();
      const selects = Array.from(document.querySelectorAll('select'));
      const exactNames = new Set([
        `variants.${variantIndex}.sizes.${sizeIndex}.stockStatus`,
        `variants[${variantIndex}].sizes[${sizeIndex}].stockStatus`,
        `variants[${variantIndex}][sizes][${sizeIndex}][stockStatus]`,
      ]);
      const stockSelect = selects.find((select) => exactNames.has(select.name))
        || (variantIndex === 0 && sizeIndex === 0 ? selects.find((select) => {
        const identity = `${select.name || ""} ${select.id || ""}`.toLowerCase();
        const label = select.id
          ? document.querySelector(`label[for="${CSS.escape(select.id)}"]`)?.innerText || ""
          : "";
        return identity.includes("stockstatus")
          || identity.includes("stock_status")
          || comparable(identity).includes("stock status")
          || comparable(label).includes("stock status")
          || Array.from(select.options).some((option) => option.value === value);
        }) : null);

      if (!stockSelect) return { found: false };

      const option = Array.from(stockSelect.options).find((item) => (
        item.value === value || comparable(item.textContent) === "out of stock"
      ));
      if (!option) return { found: true, selected: false };

      stockSelect.value = option.value;
      stockSelect.dispatchEvent(new Event('input', { bubbles: true }));
      stockSelect.dispatchEvent(new Event('change', { bubbles: true }));
      return { found: true, selected: true, selectedText: option.textContent?.trim() || value };
    }, { value: statusCode, variantIndex, sizeIndex });

    if (nativeSelectResult.selected) {
      console.log(`Selected Stock Status: "${nativeSelectResult.selectedText}" (${statusCode})`);
      return { success: true, selectedText: nativeSelectResult.selectedText };
    }
    if (nativeSelectResult.found) {
      return { success: false, reason: "option_not_found" };
    }

    const stockStatusInput = await page.evaluateHandle(({ variantIndex, sizeIndex }) => {
      const comparable = (text) => String(text || "")
        .toLowerCase()
        .replace(/[\s_-]+/g, " ")
        .trim();
      const inputs = Array.from(document.querySelectorAll('input'));
      const exactNames = new Set([
        `variants.${variantIndex}.sizes.${sizeIndex}.stockStatus`,
        `variants[${variantIndex}].sizes[${sizeIndex}].stockStatus`,
        `variants[${variantIndex}][sizes][${sizeIndex}][stockStatus]`,
      ]);

      const directMatch = inputs.find((input) => exactNames.has(input.name))
        || (variantIndex === 0 && sizeIndex === 0 ? inputs.find((input) => {
        const identity = `${input.name || ""} ${input.id || ""} ${input.placeholder || ""}`.toLowerCase();
        return identity.includes("stockstatus")
          || identity.includes("stock_status")
          || comparable(identity).includes("stock status");
        }) : null);
      if (directMatch) return directMatch;

      if (variantIndex !== 0 || sizeIndex !== 0) return null;
      for (const label of document.querySelectorAll('label')) {
        if (!comparable(label.innerText).includes("stock status")) continue;
        const labelledInput = label.htmlFor ? document.getElementById(label.htmlFor) : null;
        return labelledInput
          || label.closest('div')?.querySelector('[class*="select"] input, input')
          || label.parentElement?.querySelector('[class*="select"] input, input')
          || null;
      }

      return null;
    }, { variantIndex, sizeIndex });

    if (!(await stockStatusInput.asElement())) {
      return { success: false, reason: "input_not_found" };
    }

    return selectFromDropdown(stockStatusInput, "Out of Stock", "Stock Status", {
      exactMatch: true,
      waitForDropdown: true,
      dropdownTimeout: 3000,
      maxRetries: 3,
      optionValue: statusCode,
    });
  } catch (error) {
    console.error(`Error filling stock status: ${error.message}`);
    return { success: false, reason: "error", error: error.message };
  }
}

// ===== WARRANTY FUNCTION (WITH SELECT DROPDOWN FOR UNIT) =====
async function fillWarranty(duration, unit, description) {
  try {
    // Validate warranty unit
    const validUnits = ["days", "weeks", "months", "years", "day", "week", "month", "year"];
    let warrantyUnit = unit;
    
    if (!warrantyUnit || warrantyUnit === "" || !validUnits.includes(warrantyUnit.toLowerCase())) {
      console.log(`⚠️ Invalid warranty unit "${warrantyUnit}", using default: days`);
      warrantyUnit = "days";
    } else {
      // Normalize to plural form
      const singularMap = { "day": "days", "week": "weeks", "month": "months", "year": "years" };
      warrantyUnit = singularMap[warrantyUnit.toLowerCase()] || warrantyUnit.toLowerCase();
    }
    
    // Set default duration
    let warrantyDuration = duration;
    if (!warrantyDuration || warrantyDuration === "" || isNaN(Number(warrantyDuration))) {
      console.log(`⚠️ Invalid warranty duration "${warrantyDuration}", using default: 1`);
      warrantyDuration = 1;
    }
    
    
    // ===== FILL WARRANTY DURATION =====
    // Try to find the warranty duration input
    let durationInput = await page.$('input[name="warranty.duration"]');
    
    if (!durationInput) {
      durationInput = await page.evaluateHandle(() => {
        const allInputs = document.querySelectorAll('input[type="number"], input[type="text"]');
        for (const input of allInputs) {
          const name = (input.name || '').toLowerCase();
          const placeholder = (input.placeholder || '').toLowerCase();
          if (name.includes('duration') || 
              name.includes('warranty') && !name.includes('unit') && !name.includes('description') ||
              placeholder.includes('duration') ||
              placeholder.includes('warranty')) {
            // Check if there's a warranty label nearby
            const parent = input.closest('div');
            if (parent) {
              const label = parent.querySelector('label');
              if (label && label.innerText?.toLowerCase().includes('warranty')) {
                return input;
              }
            }
            return input;
          }
        }
        return null;
      });
    }
    
    if (durationInput && (await durationInput.asElement())) {
      await durationInput.click({ clickCount: 3 });
      await page.keyboard.press("Backspace");
      await durationInput.type(String(warrantyDuration));
    } else {
      // JavaScript fallback for duration
      await page.evaluate((duration) => {
        const allInputs = document.querySelectorAll('input');
        for (const input of allInputs) {
          const name = (input.name || '').toLowerCase();
          const placeholder = (input.placeholder || '').toLowerCase();
          if (name.includes('duration') || 
              (name.includes('warranty') && !name.includes('unit') && !name.includes('description')) ||
              placeholder.includes('duration')) {
            const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            nativeInputValueSetter.call(input, duration);
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            console.log(`✅ Filled warranty duration via JavaScript: ${duration}`);
            return true;
          }
        }
        return false;
      }, String(warrantyDuration));
    }
    
    await new Promise(r => setTimeout(r, 300));
    
    // ===== SELECT WARRANTY UNIT FROM DROPDOWN =====
    try {
      // Find the warranty unit select element
      let unitSelect = await page.$('select[name="warranty.unit"]');
      
      if (!unitSelect) {
        unitSelect = await page.evaluateHandle(() => {
          // Look for select with warranty unit
          const selects = document.querySelectorAll('select');
          for (const select of selects) {
            const name = (select.name || '').toLowerCase();
            const label = select.closest('div')?.querySelector('label')?.innerText?.toLowerCase() || '';
            if (name.includes('warranty.unit') || 
                name.includes('warranty_unit') ||
                name.includes('unit') && label.includes('warranty')) {
              return select;
            }
          }
          
          // Look by label text
          const labels = document.querySelectorAll('label');
          for (const label of labels) {
            const text = (label.innerText || '').toLowerCase();
            if (text.includes('warranty') && (text.includes('unit') || text.includes('value'))) {
              const select = label.closest('div')?.querySelector('select');
              if (select) return select;
            }
          }
          return null;
        });
      }
      
      if (unitSelect && (await unitSelect.asElement())) {
        // Select the option by value
        const selected = await unitSelect.evaluate((select, unit) => {
          const options = select.querySelectorAll('option');
          for (const option of options) {
            if (option.value.toLowerCase() === unit.toLowerCase()) {
              select.value = option.value;
              select.dispatchEvent(new Event('change', { bubbles: true }));
              select.dispatchEvent(new Event('input', { bubbles: true }));
              return true;
            }
          }
          return false;
        }, warrantyUnit);
        
        if (selected) {
        } else {
          // Try selecting by text content
          const selectedByText = await unitSelect.evaluate((select, unit) => {
            const options = select.querySelectorAll('option');
            for (const option of options) {
              if (option.textContent.toLowerCase().includes(unit.toLowerCase())) {
                select.value = option.value;
                select.dispatchEvent(new Event('change', { bubbles: true }));
                select.dispatchEvent(new Event('input', { bubbles: true }));
                return true;
              }
            }
            return false;
          }, warrantyUnit);
          
          if (selectedByText) {
            console.log(`✅ Selected warranty unit by text: ${warrantyUnit}`);
          } else {
            console.log(`⚠️ Could not find option for warranty unit: ${warrantyUnit}`);
          }
        }
      } else {
        console.log(`❌ Could not find warranty unit select dropdown`);
        
        // Try JavaScript fallback for unit
        await page.evaluate((unit) => {
          const selects = document.querySelectorAll('select');
          for (const select of selects) {
            const name = (select.name || '').toLowerCase();
            if (name.includes('warranty') && name.includes('unit')) {
              const options = select.querySelectorAll('option');
              for (const option of options) {
                if (option.value.toLowerCase() === unit.toLowerCase() ||
                    option.textContent.toLowerCase().includes(unit.toLowerCase())) {
                  select.value = option.value;
                  select.dispatchEvent(new Event('change', { bubbles: true }));
                  return true;
                }
              }
            }
          }
          return false;
        }, warrantyUnit);
        console.log(`✅ Filled warranty unit via JavaScript: ${warrantyUnit}`);
      }
    } catch (error) {
      console.error(`Error selecting warranty unit: ${error.message}`);
    }
    
    await new Promise(r => setTimeout(r, 300));
    
    // ===== FILL WARRANTY DESCRIPTION =====
    if (description && description !== "") {
      let descInput = await page.$('textarea[name="warranty.description"]');
      
      if (!descInput) {
        descInput = await page.evaluateHandle(() => {
          const textareas = document.querySelectorAll('textarea');
          for (const textarea of textareas) {
            const name = (textarea.name || '').toLowerCase();
            if (name.includes('warranty') && name.includes('description')) {
              return textarea;
            }
          }
          // Try by container/label
          const labels = document.querySelectorAll('label');
          for (const label of labels) {
            const text = (label.innerText || '').toLowerCase();
            if (text.includes('warranty') && text.includes('description')) {
              const textarea = label.closest('div')?.querySelector('textarea');
              if (textarea) return textarea;
            }
          }
          return null;
        });
      }
      
      if (descInput && (await descInput.asElement())) {
        await descInput.click({ clickCount: 3 });
        await page.keyboard.press("Backspace");
        await descInput.type(description);
        console.log(`✅ Filled warranty description`);
      } else {
        // JavaScript fallback for description
        await page.evaluate((desc) => {
          const textareas = document.querySelectorAll('textarea');
          for (const textarea of textareas) {
            const name = (textarea.name || '').toLowerCase();
            if (name.includes('warranty') && name.includes('description')) {
              const nativeTextareaValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
              nativeTextareaValueSetter.call(textarea, desc);
              textarea.dispatchEvent(new Event('input', { bubbles: true }));
              textarea.dispatchEvent(new Event('change', { bubbles: true }));
              return true;
            }
          }
          return false;
        }, description);
        console.log(`✅ Filled warranty description via JavaScript`);
      }
    }
    
    return true;
    
  } catch (error) {
    console.error(`Error filling warranty: ${error.message}`);
    return false;
  }
}

// ===== KEYWORD FUNCTION =====
async function fillKeyword(keywordValue) {
  if (!keywordValue || keywordValue === "") {
    keywordValue = "electronics";
  }
  
  try {
    // Try multiple selector strategies
    const selectors = [
      'input[name="keyword"]',
      'input[name="keywords"]',
      'input[name="tags"]',
      'input[placeholder*="keyword" i]',
      'input[placeholder*="tag" i]',
      'input[placeholder*="search term" i]',
      'input[id*="keyword" i]',
      'input[id*="tag" i]',
      // React Select style inputs
      '.css-b62m3t-container input',
      '[class*="select"] input',
      // Common form field patterns
      'input[type="text"]:not([name*="name"]):not([name*="price"])'
    ];
    
    let keywordInput = null;
    for (const selector of selectors) {
      keywordInput = await page.$(selector);
      if (keywordInput) {
        break;
      }
    }
    
    if (!keywordInput) {
      // Try evaluateHandle as fallback
      keywordInput = await page.evaluateHandle(() => {
        const allInputs = document.querySelectorAll('input[type="text"]');
        for (const input of allInputs) {
          const name = (input.name || '').toLowerCase();
          const placeholder = (input.placeholder || '').toLowerCase();
          const id = (input.id || '').toLowerCase();
          
          if (name.includes('keyword') || name.includes('tag') || 
              placeholder.includes('keyword') || placeholder.includes('tag') ||
              placeholder.includes('search') ||
              id.includes('keyword') || id.includes('tag')) {
            return input;
          }
        }
        return null;
      });
    }
    
    if (keywordInput && (await keywordInput.asElement())) {
      await keywordInput.click({ clickCount: 3 });
      await page.keyboard.press("Backspace");
      await keywordInput.type(keywordValue);
      return true;
    } else {
      // JavaScript fallback with broader search
      const filled = await page.evaluate((keyword) => {
        const allInputs = document.querySelectorAll('input[type="text"]');
        for (const input of allInputs) {
          const name = (input.name || '').toLowerCase();
          const placeholder = (input.placeholder || '').toLowerCase();
          const id = (input.id || '').toLowerCase();
          
          // Look for any field that might be for keywords/tags
          if (name.includes('keyword') || name.includes('tag') || 
              placeholder.includes('keyword') || placeholder.includes('tag') ||
              placeholder.includes('search') ||
              id.includes('keyword') || id.includes('tag')) {
            const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            nativeInputValueSetter.call(input, keyword);
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
          }
        }
        return false;
      }, keywordValue);
      
      if (filled) {
        console.log(`✅ Filled keyword via JavaScript: ${keywordValue}`);
        return true;
      } else {
        console.log(`⚠️ Could not find keyword field, skipping...`);
        return false;
      }
    }
  } catch (error) {
    console.error(`Error filling keyword: ${error.message}`);
    return false;
  }
}
async function downloadImage(url, index, productId) {
  const result = await downloadImagesFromSource(url, index, productId);
  return result.paths[0] || null;
}

const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".avif"]);
const IMAGE_CONTENT_TYPE_EXTENSIONS = {
  "image/jpeg": ".jpg",
  "image/jpg": ".jpg",
  "image/png": ".png",
  "image/webp": ".webp",
  "image/gif": ".gif",
  "image/bmp": ".bmp",
  "image/avif": ".avif",
};

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function sanitizeFilePart(value) {
  const safeValue = cleanText(value).replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").replace(/\s+/g, "_");
  return safeValue.slice(0, 80) || "product";
}

function uniqueValues(values) {
  return [...new Set(values.filter(Boolean))];
}

function uniqueImageUrlsByFileName(urls) {
  const seen = new Set();
  const uniqueUrls = [];

  for (const url of urls) {
    const parsedUrl = getUrlObject(url);
    const fileName = parsedUrl ? path.basename(parsedUrl.pathname).toLowerCase() : url;
    const key = fileName || url;

    if (!seen.has(key)) {
      seen.add(key);
      uniqueUrls.push(url);
    }
  }

  return uniqueUrls;
}

function decodeHtml(value) {
  return cleanText(value)
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function toAbsoluteUrl(value, baseUrl) {
  try {
    const decoded = decodeHtml(value);
    if (decoded.startsWith("//")) {
      return `https:${decoded}`;
    }
    return new URL(decoded, baseUrl).href;
  } catch {
    return null;
  }
}

function getUrlObject(value) {
  try {
    return new URL(value);
  } catch {
    return null;
  }
}

function isImageExtension(url) {
  const parsedUrl = getUrlObject(url);
  if (!parsedUrl) return false;
  return IMAGE_EXTENSIONS.has(path.extname(parsedUrl.pathname).toLowerCase());
}

function isPostimgHostedImage(url) {
  const parsedUrl = getUrlObject(url);
  return parsedUrl?.hostname === "i.postimg.cc";
}

function isPostimgPage(url) {
  const parsedUrl = getUrlObject(url);
  return parsedUrl?.hostname === "postimg.cc" || parsedUrl?.hostname.endsWith(".postimg.cc");
}

function getExtensionFromUrlOrType(url, contentType) {
  const parsedUrl = getUrlObject(url);
  const urlExtension = parsedUrl ? path.extname(parsedUrl.pathname).toLowerCase() : "";
  if (IMAGE_EXTENSIONS.has(urlExtension)) {
    return urlExtension;
  }

  const normalizedContentType = cleanText(contentType).split(";")[0].toLowerCase();
  return IMAGE_CONTENT_TYPE_EXTENSIONS[normalizedContentType] || ".jpg";
}

async function getWithRetries(url, options = {}, retries = 3) {
  let lastError;

  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      return await axios({
        method: "GET",
        url,
        responseType: "arraybuffer",
        timeout: 45000,
        maxRedirects: 5,
        validateStatus: status => status >= 200 && status < 400,
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
          "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
          ...options.headers,
        },
        ...options,
      });
    } catch (error) {
      lastError = error;
      if (attempt < retries) {
        await delay(1000 * attempt);
      }
    }
  }

  throw lastError;
}

function extractImageUrlsFromHtml(html, baseUrl) {
  const candidates = [];
  const metaPatterns = [
    /<meta[^>]+(?:property|name)=["'](?:og:image|og:image:secure_url|twitter:image|twitter:image:src)["'][^>]+content=["']([^"']+)["'][^>]*>/gi,
    /<meta[^>]+content=["']([^"']+)["'][^>]+(?:property|name)=["'](?:og:image|og:image:secure_url|twitter:image|twitter:image:src)["'][^>]*>/gi,
  ];

  for (const pattern of metaPatterns) {
    for (const match of html.matchAll(pattern)) {
      candidates.push(toAbsoluteUrl(match[1], baseUrl));
    }
  }

  for (const match of html.matchAll(/https?:\/\/i\.postimg\.cc\/[^\s"'<>\\)]+/gi)) {
    candidates.push(toAbsoluteUrl(match[0], baseUrl));
  }

  for (const match of html.matchAll(/\/\/i\.postimg\.cc\/[^\s"'<>\\)]+/gi)) {
    candidates.push(toAbsoluteUrl(match[0], baseUrl));
  }

  return uniqueImageUrlsByFileName(uniqueValues(
    candidates.filter(url => url && (isPostimgHostedImage(url) || isImageExtension(url)))
  ));
}

function extractPostimgPageUrlsFromHtml(html, baseUrl) {
  const links = [];

  for (const match of html.matchAll(/(?:href|data-url)=["']([^"']+)["']/gi)) {
    const absoluteUrl = toAbsoluteUrl(match[1], baseUrl);
    const parsedUrl = getUrlObject(absoluteUrl);
    if (!parsedUrl) continue;

    const isImagePage = parsedUrl.hostname === "postimg.cc" &&
      !parsedUrl.pathname.startsWith("/gallery/") &&
      /^\/[A-Za-z0-9]+\/?$/.test(parsedUrl.pathname);

    if (isImagePage) {
      links.push(parsedUrl.href);
    }
  }

  return uniqueValues(links);
}

async function resolveImageUrls(sourceUrl, depth = 0) {
  if (isImageExtension(sourceUrl) || isPostimgHostedImage(sourceUrl)) {
    return [sourceUrl];
  }

  const response = await getWithRetries(sourceUrl);
  const contentType = cleanText(response.headers["content-type"]).toLowerCase();

  if (contentType.startsWith("image/")) {
    return [sourceUrl];
  }

  const html = Buffer.from(response.data).toString("utf8");
  const directImageUrls = extractImageUrlsFromHtml(html, sourceUrl);
  if (directImageUrls.length > 0) {
    return directImageUrls;
  }

  if (depth >= 1 || !isPostimgPage(sourceUrl)) {
    throw new Error(`Could not find a direct image inside ${sourceUrl}`);
  }

  const imagePageUrls = extractPostimgPageUrlsFromHtml(html, sourceUrl).slice(0, MAX_IMAGES_PER_PRODUCT);
  const resolvedUrls = [];

  for (const imagePageUrl of imagePageUrls) {
    try {
      resolvedUrls.push(...await resolveImageUrls(imagePageUrl, depth + 1));
    } catch (error) {
      console.log(`Could not resolve gallery image ${imagePageUrl}: ${error.message}`);
    }
  }

  return uniqueValues(resolvedUrls);
}

async function downloadRemoteImage(imageUrl, sourceIndex, imageIndex, productId) {
  const response = await getWithRetries(imageUrl);
  const contentType = cleanText(response.headers["content-type"]).toLowerCase();

  if (!contentType.startsWith("image/")) {
    throw new Error(`Downloaded content is not an image (${contentType || "unknown type"})`);
  }

  const extension = getExtensionFromUrlOrType(imageUrl, contentType);
  const safeProductId = sanitizeFilePart(productId);
  const imagePath = path.join(DOWNLOAD_DIR, `${safeProductId}_image_${sourceIndex + 1}_${imageIndex + 1}${extension}`);

  try {
    fs.writeFileSync(imagePath, Buffer.from(response.data));

    const stats = fs.statSync(imagePath);
    if (stats.size === 0) {
      throw new Error(`Downloaded image is empty: ${imageUrl}`);
    }
  } catch (error) {
    if (fs.existsSync(imagePath)) {
      fs.unlinkSync(imagePath);
    }
    throw error;
  }

  const uploadPath = await prepareApiUploadImage(imagePath);
  if (uploadPath !== imagePath) {
    console.log(`Converted AVIF image to JPEG for IBSHER: ${path.basename(uploadPath)}`);
  }
  return uploadPath;
}

function copyLocalImage(localImagePath, sourceIndex, imageIndex, productId) {
  const extension = path.extname(localImagePath).toLowerCase();
  const safeProductId = sanitizeFilePart(productId);
  const copiedImagePath = path.join(
    DOWNLOAD_DIR,
    `${safeProductId}_local_${sourceIndex + 1}_${imageIndex + 1}${extension}`,
  );

  try {
    fs.copyFileSync(localImagePath, copiedImagePath);
    const stats = fs.statSync(copiedImagePath);
    if (!stats.isFile() || stats.size === 0) {
      throw new Error(`Local image is empty: ${localImagePath}`);
    }
  } catch (error) {
    if (fs.existsSync(copiedImagePath)) fs.unlinkSync(copiedImagePath);
    throw error;
  }

  return copiedImagePath;
}

async function downloadImagesFromSource(sourceUrl, sourceIndex, productId) {
  const result = {
    paths: [],
    failures: [],
  };

  try {
    const localSource = resolveLocalImageSource(sourceUrl, {
      imageFolder: IMAGE_SOURCE_DIR,
      workbookFile: SOURCE_EXCEL_FILE,
      imageExtensions: IMAGE_EXTENSIONS,
    });

    if (localSource.handled) {
      if (localSource.error) throw new Error(localSource.error);

      const localPaths = localSource.paths.slice(0, MAX_IMAGES_PER_PRODUCT);
      console.log(`Preparing ${localPaths.length} local image(s) from source ${sourceIndex + 1}: ${sourceUrl}`);

      for (let i = 0; i < localPaths.length; i++) {
        try {
          const copiedPath = copyLocalImage(localPaths[i], sourceIndex, i, productId);
          const uploadPath = await prepareApiUploadImage(copiedPath);
          if (uploadPath !== copiedPath) {
            console.log(`Converted AVIF image to JPEG for IBSHER: ${path.basename(uploadPath)}`);
          }
          result.paths.push(uploadPath);
        } catch (error) {
          result.failures.push({
            url: localPaths[i],
            reason: error.message,
          });
          console.log(`Skipping failed local image ${localPaths[i]}: ${error.message}`);
        }
      }

      console.log(`Prepared ${result.paths.length} local image(s) from source ${sourceIndex + 1}`);
      return result;
    }

    console.log(`Downloading image source ${sourceIndex + 1}: ${sourceUrl}`);
    const resolvedUrls = (await resolveImageUrls(sourceUrl)).slice(0, MAX_IMAGES_PER_PRODUCT);

    if (resolvedUrls.length === 0) {
      throw new Error(`No downloadable images found at ${sourceUrl}`);
    }

    for (let i = 0; i < resolvedUrls.length; i++) {
      try {
        const imagePath = await downloadRemoteImage(resolvedUrls[i], sourceIndex, i, productId);
        result.paths.push(imagePath);
      } catch (error) {
        result.failures.push({
          url: resolvedUrls[i],
          reason: error.message,
        });
        console.log(`Skipping failed image ${resolvedUrls[i]}: ${error.message}`);
      }
    }

    console.log(`Downloaded ${result.paths.length} image(s) from source ${sourceIndex + 1}`);
  } catch (error) {
    result.failures.push({
      url: sourceUrl,
      reason: `Failed to download image source ${sourceIndex + 1}: ${error.message}`,
    });
    console.log(`Skipping failed image source ${sourceIndex + 1}: ${error.message}`);
  }

  return result;
}

function trackImageDownloadFailures(product, sourceUrl, sourceIndex, failures = [], variant = null) {
  failures.forEach((failure) => {
    failedItems.imageDownloadFailures.push({
      itemcode: product.itemcode,
      name: product.en,
      color: cleanText(variant?.color),
      imageNumber: sourceIndex + 1,
      imageUrl: failure.url || sourceUrl,
      reason: `${variant?.color ? `${variant.color}: ` : ""}${failure.reason || "Failed to download image"}`,
    });
  });
}

async function findImageFileInput(variantIndex = 0) {
  const variantImageSelectors = [
    `input[type="file"][name="variants.${variantIndex}.images"]`,
    `input[type="file"][name="variants[${variantIndex}].images"]`,
    `input[type="file"][name^="variants.${variantIndex}."][name*=".images"]`,
    `input[type="file"][name^="variants[${variantIndex}]"][name*="images"]`,
    `[data-variant-index="${variantIndex}"] input[type="file"]`,
  ];

  if (variantIndex === 0) {
    variantImageSelectors.push(
      '.variant-image-input input[type="file"]',
      'input[type="file"][accept*="image"]',
      'input[type="file"]',
    );
  }

  for (const selector of variantImageSelectors) {
    try {
      const fileInput = await page.waitForSelector(selector, { timeout: 3000 });
      if (fileInput) {
        return fileInput;
      }
    } catch {
      // Try the next selector.
    }
  }

  return null;
}

async function findClickableByText(textValues) {
  const handle = await page.evaluateHandle((values) => {
    const normalizedValues = values.map(value => value.toLowerCase());
    const candidates = [...document.querySelectorAll('button, [role="button"], label, a')];

    return candidates.find(element => {
      const rect = element.getBoundingClientRect();
      const style = window.getComputedStyle(element);
      if (rect.width === 0 || rect.height === 0 || style.display === "none" || style.visibility === "hidden") {
        return false;
      }

      const text = (element.innerText || element.textContent || element.getAttribute("aria-label") || "").trim().toLowerCase();
      return normalizedValues.some(value => text === value || text.includes(value));
    }) || null;
  }, textValues);

  return handle.asElement();
}

async function findVariantActionButton(variantIndex, textValues) {
  const handle = await page.evaluateHandle(({ variantIndex, textValues }) => {
    const normalizedValues = textValues.map(value => value.toLowerCase());
    const isVisible = (element) => {
      const rect = element.getBoundingClientRect();
      const style = window.getComputedStyle(element);
      return rect.width > 0 && rect.height > 0 && style.display !== "none" && style.visibility !== "hidden";
    };
    const matches = (element) => {
      const text = (element.innerText || element.textContent || element.getAttribute("aria-label") || "")
        .trim()
        .toLowerCase();
      return isVisible(element) && normalizedValues.some(value => text === value || text.includes(value));
    };
    const colorInput = document.querySelector(`input[name="variants.${variantIndex}.color.en"]`)
      || document.querySelector(`input[name="variants[${variantIndex}].color.en"]`);
    let container = colorInput?.parentElement || null;

    for (let depth = 0; container && depth < 8; depth += 1, container = container.parentElement) {
      const scoped = [...container.querySelectorAll('button, [role="button"]')].find(matches);
      if (scoped) return scoped;
    }

    return [...document.querySelectorAll('button, [role="button"]')].find(matches) || null;
  }, { variantIndex, textValues });

  return handle.asElement();
}

async function ensureVariantOption(variantIndex) {
  if (variantIndex === 0) return;
  const selector = `input[name="variants.${variantIndex}.color.en"], input[name="variants[${variantIndex}].color.en"]`;
  if (await page.$(selector)) return;

  const button = await findClickableByText(["add option", "add variant"]);
  if (!button) throw new Error(`Could not find Add Option for color option ${variantIndex + 1}`);
  await button.click();
  await page.waitForSelector(selector, { visible: true, timeout: 5000 });
}

async function ensureVariantSize(variantIndex, sizeIndex) {
  if (sizeIndex === 0) return;
  const selector = `input[name="variants.${variantIndex}.sizes.${sizeIndex}.size"], input[name="variants[${variantIndex}].sizes[${sizeIndex}].size"]`;
  if (await page.$(selector)) return;

  const button = await findVariantActionButton(variantIndex, ["add size"]);
  if (!button) throw new Error(`Could not find Add Size for color option ${variantIndex + 1}`);
  await button.click();
  await page.waitForSelector(selector, { visible: true, timeout: 5000 });
}

async function clickAddImageButton(variantIndex = 0) {
  const selectors = [
    `[data-variant-index="${variantIndex}"] .add-image`,
    `[data-variant-index="${variantIndex}"] .btn-add-image`,
    `[data-variant-index="${variantIndex}"] [class*="add"][class*="image"]`,
    `[data-variant-index="${variantIndex}"] [class*="add"][class*="photo"]`,
  ];

  for (const selector of selectors) {
    try {
      const element = await page.$(selector);
      if (element) {
        await element.click();
        await delay(1000);
        return true;
      }
    } catch {
      // Try text-based lookup below.
    }
  }

  const button = await findVariantActionButton(
    variantIndex,
    ["add image", "add photo", "add more image", "add more photo"],
  );
  if (!button) {
    return false;
  }

  await button.click();
  await delay(1000);
  return true;
}

async function getUploadState(filePaths = []) {
  const fileNames = filePaths.map(filePath => path.basename(filePath));

  return page.evaluate((names) => {
    const fileInputs = [...document.querySelectorAll('input[type="file"]')];
    const inputFileCount = fileInputs.reduce((sum, input) => sum + (input.files?.length || 0), 0);
    const bodyText = document.body?.innerText || "";
    const fileNameMatches = names.filter(name => bodyText.includes(name)).length;
    const visiblePreviewCount = [...document.querySelectorAll("img")].filter(img => {
      const rect = img.getBoundingClientRect();
      const style = window.getComputedStyle(img);
      const src = img.currentSrc || img.src || "";
      return rect.width > 0 &&
        rect.height > 0 &&
        style.display !== "none" &&
        style.visibility !== "hidden" &&
        (src.startsWith("blob:") || src.startsWith("data:image") || src.startsWith("http"));
    }).length;

    return { inputFileCount, fileNameMatches, visiblePreviewCount };
  }, fileNames);
}

function uploadLooksSuccessful(currentState, beforeState, expectedCount) {
  return currentState.inputFileCount >= expectedCount ||
    currentState.fileNameMatches >= expectedCount ||
    currentState.visiblePreviewCount >= beforeState.visiblePreviewCount + expectedCount;
}

async function waitForUploadConfirmation(filePaths, beforeState) {
  const expectedCount = filePaths.length;

  await page.waitForFunction(
    ({ expectedCount, beforeVisiblePreviewCount, fileNames }) => {
      const fileInputs = [...document.querySelectorAll('input[type="file"]')];
      const inputFileCount = fileInputs.reduce((sum, input) => sum + (input.files?.length || 0), 0);
      const bodyText = document.body?.innerText || "";
      const fileNameMatches = fileNames.filter(name => bodyText.includes(name)).length;
      const visiblePreviewCount = [...document.querySelectorAll("img")].filter(img => {
        const rect = img.getBoundingClientRect();
        const style = window.getComputedStyle(img);
        const src = img.currentSrc || img.src || "";
        return rect.width > 0 &&
          rect.height > 0 &&
          style.display !== "none" &&
          style.visibility !== "hidden" &&
          (src.startsWith("blob:") || src.startsWith("data:image") || src.startsWith("http"));
      }).length;

      return inputFileCount >= expectedCount ||
        fileNameMatches >= expectedCount ||
        visiblePreviewCount >= beforeVisiblePreviewCount + expectedCount;
    },
    { timeout: 30000 },
    {
      expectedCount,
      beforeVisiblePreviewCount: beforeState.visiblePreviewCount,
      fileNames: filePaths.map(filePath => path.basename(filePath)),
    }
  );
}

async function uploadVariantImages(product, variant, variantIndex) {
  const imageUrls = variant.images;
  if (!imageUrls || imageUrls.length === 0) {
    console.log("No images to upload");
    return [];
  }

  const downloadedPaths = [];
  const uploadFailures = [];
  // Download all images first
  for (let i = 0; i < imageUrls.length; i++) {
    const result = await downloadImagesFromSource(
      imageUrls[i],
      i,
      `${product.itemcode}_${variant.color || variantIndex + 1}`,
    );
    downloadedPaths.push(...result.paths);
    trackImageDownloadFailures(product, imageUrls[i], i, result.failures, variant);
  }

  const limitedDownloadedPaths = downloadedPaths.slice(0, MAX_IMAGES_PER_PRODUCT);

  if (limitedDownloadedPaths.length === 0) {
    console.log(`No valid images were downloaded for ${product.itemcode}; continuing without images`);
    return [];
  }

  let fileInput = await findImageFileInput(variantIndex);

  if (!fileInput) {
    throw new Error("No image file input found on page");
  }

  const beforeUploadState = await getUploadState(limitedDownloadedPaths);
  const supportsMultiple = await fileInput.evaluate(input => Boolean(input.multiple));

  if (supportsMultiple) {
    await fileInput.uploadFile(...limitedDownloadedPaths);
  } else {
    for (let i = 0; i < limitedDownloadedPaths.length; i++) {
      try {
        await fileInput.uploadFile(limitedDownloadedPaths[i]);
        await delay(1500);

        if (i < limitedDownloadedPaths.length - 1) {
          const addedInput = await clickAddImageButton(variantIndex);
          if (!addedInput) {
            throw new Error("Image input does not allow multiple files and no Add Image button was found");
          }
          fileInput = await findImageFileInput(variantIndex);
        }
      } catch (error) {
        uploadFailures.push(`Image ${i + 1}: ${error.message}`);
      }
    }
  }

  if (uploadFailures.length > 0) {
    throw new Error(`Image upload failed: ${uploadFailures.join("; ")}`);
  }

  await waitForUploadConfirmation(limitedDownloadedPaths, beforeUploadState);
  const afterUploadState = await getUploadState(limitedDownloadedPaths);

  if (!uploadLooksSuccessful(afterUploadState, beforeUploadState, limitedDownloadedPaths.length)) {
    throw new Error("Images were selected but the page did not confirm that they were attached");
  }

  console.log(`Uploaded ${limitedDownloadedPaths.length} image(s) successfully`);
  return limitedDownloadedPaths;
}

function cleanupProductImages(imagePaths) {
  if (imagePaths && imagePaths.length > 0) {
    for (const imagePath of imagePaths) {
      try {
        if (fs.existsSync(imagePath)) {
          fs.unlinkSync(imagePath);
        }
      } catch (error) {
        console.error(`Failed to delete ${imagePath}:`, error.message);
      }
    }
  }
}

// ===== API AUTOMATION =====
const apiCookies = new Map();
const apiClient = axios.create({
  baseURL: API_BASE_URL.replace(/\/+$/, ""),
  timeout: 60000,
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
  headers: { Accept: "application/json" },
});
let apiRefreshPromise = null;

function captureApiCookies(headers = {}) {
  const setCookieHeaders = headers["set-cookie"] || headers["Set-Cookie"] || [];
  const values = Array.isArray(setCookieHeaders) ? setCookieHeaders : [setCookieHeaders];

  for (const value of values) {
    const firstPart = cleanText(value).split(";", 1)[0];
    const separator = firstPart.indexOf("=");
    if (separator <= 0) continue;

    const name = firstPart.slice(0, separator).trim();
    const cookieValue = firstPart.slice(separator + 1).trim();
    if (cookieValue) apiCookies.set(name, cookieValue);
    else apiCookies.delete(name);
  }
}

function getApiCookieHeader() {
  return [...apiCookies.entries()].map(([name, value]) => `${name}=${value}`).join("; ");
}

function withApiCookies(config = {}) {
  const cookie = getApiCookieHeader();
  return {
    ...config,
    headers: {
      ...(config.headers || {}),
      ...(cookie ? { Cookie: cookie } : {}),
    },
  };
}

async function refreshApiSession() {
  if (!apiRefreshPromise) {
    apiRefreshPromise = (async () => {
      const response = await apiClient.request(withApiCookies({
        method: "POST",
        url: "/user/refresh-token",
        data: {},
      }));
      captureApiCookies(response.headers);
      return response;
    })().finally(() => {
      apiRefreshPromise = null;
    });
  }

  return apiRefreshPromise;
}

async function apiRequest(config, { refreshOnUnauthorized = true } = {}) {
  try {
    const response = await apiClient.request(withApiCookies(config));
    captureApiCookies(response.headers);
    return response;
  } catch (error) {
    if (refreshOnUnauthorized && error.response?.status === 401) {
      await refreshApiSession();
      const response = await apiClient.request(withApiCookies(config));
      captureApiCookies(response.headers);
      return response;
    }
    throw error;
  }
}

async function loginToIbsherApi(credentials) {
  apiCookies.clear();
  const response = await apiClient.post("/user/login", {
    email: credentials.email,
    password: credentials.password,
  });
  captureApiCookies(response.headers);

  if (apiCookies.size === 0) {
    throw new Error("IBSHER API login succeeded but did not return an authentication cookie");
  }

  return response.data;
}

function getApiErrorMessage(error) {
  const data = error?.response?.data;
  const status = error?.response?.status;
  let message = "";
  const fieldErrors = data?.errors?.fieldErrors;
  if (fieldErrors && typeof fieldErrors === "object") {
    message = Object.entries(fieldErrors)
      .flatMap(([field, values]) => (Array.isArray(values) ? values : [values])
        .map(value => `${field}: ${value}`)).join("; ");
  }
  if (Array.isArray(data?.errors?.formErrors) && data.errors.formErrors.length > 0) {
    message = [message, ...data.errors.formErrors].filter(Boolean).join("; ");
  }
  if (!message) {
    message = typeof data === "string"
      ? cleanText(data)
      : cleanText(data?.message || data?.error || (data ? JSON.stringify(data) : ""));
  }
  message ||= error?.message || "Unknown API error";
  const requestId = cleanText(error?.response?.headers?.["x-request-id"]);
  return `${status ? `HTTP ${status}: ` : ""}${message}${requestId ? ` (request ID: ${requestId})` : ""}`;
}

function normalizeLookupKey(value) {
  return cleanText(value).toLocaleLowerCase().replace(/\s+/g, " ");
}

function buildEntityLookup(items = []) {
  const lookup = new Map();

  for (const item of items) {
    const id = cleanText(item?._id || item?.id);
    if (!id) continue;
    const names = item?.name && typeof item.name === "object"
      ? Object.values(item.name)
      : [item?.name, item?.label];

    for (const value of [id, ...names]) {
      const key = normalizeLookupKey(value);
      if (key && !lookup.has(key)) lookup.set(key, id);
    }
  }

  return lookup;
}

function getProductKeywords(product) {
  const values = Array.isArray(product.keyword)
    ? product.keyword
    : cleanText(product.keyword).split(/[,;\n]+/);
  const keywords = uniqueValues(values.map(cleanText));
  return keywords.length > 0
    ? keywords
    : [product.en.split(/\s+/).slice(0, 3).join(" ").toLowerCase()].filter(Boolean);
}

function buildApiProductForm(product, brandId, collection, variantImagePaths) {
  const form = new FormData();

  form.append("name", JSON.stringify({ en: product.en, ar: product.ar, ku: product.ku }));
  form.append("description", JSON.stringify({
    en: product.descriptionEn,
    ar: product.descriptionAr,
    ku: product.descriptionKu,
  }));
  form.append(IBSHER_ITEM_CODE_FIELD, product.itemcode);
  form.append("keyword", JSON.stringify(getProductKeywords(product)));
  form.append("category", collection.categoryId);
  form.append("subCategory", collection.subCategoryId);
  form.append("collectionName", collection.id);
  form.append("brand", brandId);
  // Use the same numeric conversion as Excel validation. Comma-formatted
  // values must not pass validation and then reach the API as numeric strings
  // containing commas (or become null inside warranty JSON).
  form.append("currency", product.currency === "USD" ? "USD" : "IQD");
  form.append("price", String(parseNumber(product.price)));
  form.append("discountPrice", String(isMissing(product.discountPrice) ? 0 : parseNumber(product.discountPrice)));
  form.append("is_featured", "false");
  form.append("is_new_arrival", "false");
  form.append("is_hot", "false");
  form.append("is_best_seller", "false");
  form.append("badge", "");
  form.append("rating", "0");
  form.append("isActive", "true");

  if (product.warrantyDuration || product.warrantyUnit || product.warrantyDescription) {
    form.append("warranty", JSON.stringify({
      duration: isMissing(product.warrantyDuration) ? undefined : parseNumber(product.warrantyDuration),
      unit: product.warrantyUnit || undefined,
      description: product.warrantyDescription || undefined,
    }));
  }

  form.append("variants", JSON.stringify(buildApiVariants(product)));

  variantImagePaths.forEach((imagePaths, variantIndex) => {
    imagePaths.forEach((imagePath, imageIndex) => {
      form.append(
        buildVariantImageFieldName(variantIndex, imageIndex),
        fs.createReadStream(imagePath),
      );
    });
  });

  return form;
}

async function submitApiProduct(product, brandId, collection, variantImagePaths) {
  let lastError;

  for (let attempt = 1; attempt <= 3; attempt++) {
    const form = buildApiProductForm(product, brandId, collection, variantImagePaths);
    try {
      return await apiRequest({
        method: "POST",
        url: "/product",
        data: form,
        headers: form.getHeaders(),
      }, { refreshOnUnauthorized: false });
    } catch (error) {
      lastError = error;
      if (error.response?.status === 401 && attempt < 3) {
        await refreshApiSession();
        continue;
      }
      if (error.response?.status === 429 && attempt < 3) {
        await delay(750 * attempt);
        continue;
      }
      throw error;
    }
  }

  throw lastError;
}

async function processApiProduct(product, productIndex, lookups) {
  const downloadedVariantImages = product.variants.map(() => []);
  console.log(`\n[API] Product ${productIndex + 1}/${products.length}: ${product.itemcode} - ${product.en}`);

  const brandId = lookups.brands.get(normalizeLookupKey(product.brand));
  const collection = resolveIbsherCollection(product.collectionName, lookups.catalog);
  const collectionHierarchyFound = Boolean(
    collection
    && !collection.ambiguous
    && collection.categoryId
    && collection.subCategoryId
    && collection.id
  );
  if (!brandId || !collectionHierarchyFound) {
    if (!brandId) {
      failedItems.dropdownNotFound.push({
        itemcode: product.itemcode,
        name: product.en,
        field: "brand",
        value: product.brand,
        reason: "Brand was not found by the API",
      });
    }
    if (!collectionHierarchyFound) {
      failedItems.dropdownNotFound.push({
        itemcode: product.itemcode,
        name: product.en,
        field: "collection",
        value: product.collectionName,
        reason: collection?.ambiguous
          ? "Collection name matches more than one IBSHER category path"
          : "Collection or its category hierarchy was not found by the API",
      });
    }
    console.log(`[API] Skipped ${product.itemcode}: brand or collection was not found`);
    return;
  }

  try {
    const imageLimitPerVariant = Math.min(MAX_IMAGES_PER_PRODUCT, 10);
    for (let variantIndex = 0; variantIndex < product.variants.length; variantIndex++) {
      const variant = product.variants[variantIndex];
      const variantImages = downloadedVariantImages[variantIndex];
      for (let imageIndex = 0; imageIndex < variant.images.length; imageIndex++) {
        if (variantImages.length >= imageLimitPerVariant) break;
        const result = await downloadImagesFromSource(
          variant.images[imageIndex],
          imageIndex,
          `${product.itemcode}_${variant.color || variantIndex + 1}`,
        );
        variantImages.push(...result.paths);
        trackImageDownloadFailures(
          product,
          variant.images[imageIndex],
          imageIndex,
          result.failures,
          variant,
        );
      }
      downloadedVariantImages[variantIndex] = variantImages.slice(0, imageLimitPerVariant);
    }

    await submitApiProduct(product, brandId, collection, downloadedVariantImages);
    processedItemcodes.add(product.itemcode);
    const downloadedImageCount = downloadedVariantImages.flat().length;
    console.log(
      `[API] Created ${product.itemcode} successfully with ${product.variants.length} color option(s)`
      + `${downloadedImageCount ? ` and ${downloadedImageCount} image(s)` : ""}`,
    );
  } catch (error) {
    const reason = getApiErrorMessage(error);
    failedItems.otherErrors.push({ itemcode: product.itemcode, name: product.en, reason });
    console.error(`[API] Failed ${product.itemcode}: ${reason}`);
    if (error.response?.status >= 500) {
      console.error(
        `[API] ${product.itemcode} request details: price=${parseNumber(product.price)},`
        + ` discountPrice=${isMissing(product.discountPrice) ? 0 : parseNumber(product.discountPrice)},`
        + ` color options=${product.variants.length}, images per option=[${downloadedVariantImages.map(images => images.length).join(", ")}]`,
      );
    }
  } finally {
    cleanupProductImages(downloadedVariantImages.flat());
  }
}

async function fetchIbsherApiListPages(url, keys = [], label = url) {
  return collectIbsherApiListPages({
    keys,
    limit: 100,
    incompleteMessage: `Could not load the complete IBSHER ${label} list.`,
    tooLargeMessage: `The IBSHER ${label} list is too large to load completely.`,
    requestPage: async (page, limit) => {
      const response = await apiRequest({ method: "GET", url, params: { page, limit } });
      return response.data;
    },
  });
}

async function loadIbsherApiLookups(credentials) {
  await loginToIbsherApi(credentials);
  console.log("[API] Login successful. Loading brands and the category hierarchy...");

  const [brands, categoryResponse, subCategories, collections] = await Promise.all([
    fetchIbsherApiListPages("/brand", ["brands", "brand", "data"], "brand"),
    apiRequest({ method: "GET", url: "/category" }),
    fetchIbsherApiListPages("/subcategory", [
      "subCategory", "subCategories", "subcategory", "subcategories", "data",
    ], "subcategory"),
    fetchIbsherApiListPages("/collection", ["collections", "collection", "data"], "collection"),
  ]);
  const lookups = {
    brands: buildEntityLookup(brands),
    catalog: buildIbsherProductCatalog(
      categoryResponse.data,
      { subCategory: subCategories },
      { collections }
    ),
  };

  console.log(
    `[API] Loaded ${lookups.brands.size} brand key(s), `
    + `${lookups.catalog.categories.length} categories, `
    + `${lookups.catalog.subCategories.length} subcategories, and `
    + `${lookups.catalog.collections.length} collections.`
  );
  return lookups;
}

async function runApiAutomation(credentials) {
  console.log(`Starting API automation at ${API_BASE_URL}`);
  const lookups = await loadIbsherApiLookups(credentials);
  console.log(`[API] Uploading with concurrency ${API_CONCURRENCY}; API image limit is 10 per color option.`);

  let nextProductIndex = 0;
  async function worker() {
    while (nextProductIndex < products.length) {
      const productIndex = nextProductIndex++;
      await processApiProduct(products[productIndex], productIndex, lookups);
    }
  }

  await Promise.all(Array.from({ length: Math.min(API_CONCURRENCY, products.length) }, () => worker()));
}

// ===== GET USER CHOICE =====
submissionMode = AUTOMATION_MODE === "api" ? "auto" : await getUserChoice();
const ibsherCredentials = await getIbsherCredentials();
console.log(`\nAutomation engine: ${AUTOMATION_MODE === "api" ? "API (fast)" : "Chrome"}`);
console.log(`Selected submission mode: ${submissionMode === 'auto' ? 'Auto Submit' : 'Manual Submit'}`);

if (AUTOMATION_MODE === "api") {
  await runApiAutomation(ibsherCredentials);
} else {
  console.log("Starting browser automation...\n");
  const browserLookups = await loadIbsherApiLookups(ibsherCredentials);
  const browserCatalog = browserLookups.catalog;


// ===== LOGIN =====
await page.goto("https://marketplace.ibsher.com", {
  waitUntil: "networkidle2",
  timeout: 30000
});

await page.waitForSelector('input[type="email"]', { visible: true, timeout: 10000 });

await page.type("input[type='email']", ibsherCredentials.email);
await page.type("input[type='password']", ibsherCredentials.password);

await Promise.all([
  page.click('button[type="submit"]'),
  page.waitForNavigation({ waitUntil: "networkidle2", timeout: 30000 }),
]);

// ===== GO TO ADD PRODUCT PAGE =====
await page.goto("https://marketplace.ibsher.com/add-product", {
  waitUntil: "networkidle2",
  timeout: 30000
});

// ===== LOOP PRODUCTS =====
for (let productIndex = 0; productIndex < products.length; productIndex++) {
  const product = products[productIndex];
  let downloadedImages = [];
  let skipProduct = false;
  let dropdownFailed = false;
  
  console.log(`\n${"=".repeat(60)}`);
  console.log(`📦 PRODUCT ${productIndex + 1}/${products.length}`);
  console.log(`${"=".repeat(60)}`);
  console.log(`Item Code: ${product.itemcode}`);
  console.log(`Name: ${product.en}`);

  // ===== CHECK FOR DUPLICATE ITEM CODE =====
  if (processedItemcodes.has(product.itemcode)) {
    console.log(`\n⚠️ DUPLICATE ITEM CODE DETECTED: ${product.itemcode}`);
    console.log(`   This item code has already been processed. Skipping...`);
    failedItems.duplicateItemcode.push({
      itemcode: product.itemcode,
      name: product.en,
      reason: "Duplicate item code"
    });
    continue;
  }

  // Check if brand and collection exist
  const hasBrand = product.brand && product.brand !== "";
  const hasCollection = product.collectionName && product.collectionName !== "";
  
  if (!hasBrand && !hasCollection) {
    console.log(`\n❌ SKIPPING: Item ${product.itemcode} has NO BRAND and NO COLLECTION`);
    console.log(`   This item will not be added to the website.`);
    failedItems.noBrandAndCollection.push({
      itemcode: product.itemcode,
      name: product.en,
      brand: product.brand,
      collection: product.collectionName,
      reason: "Missing both brand and collection"
    });
    skipProduct = true;
    continue;
  } else if (!hasBrand) {
    console.log(`\n❌ SKIPPING: Item ${product.itemcode} has NO BRAND`);
    console.log(`   This item will not be added to the website.`);
    failedItems.noBrand.push({
      itemcode: product.itemcode,
      name: product.en,
      brand: product.brand,
      collection: product.collectionName,
      reason: "Missing brand"
    });
    skipProduct = true;
    continue;
  } else if (!hasCollection) {
    console.log(`\n❌ SKIPPING: Item ${product.itemcode} has NO COLLECTION`);
    console.log(`   This item will not be added to the website.`);
    failedItems.noCollection.push({
      itemcode: product.itemcode,
      name: product.en,
      brand: product.brand,
      collection: product.collectionName,
      reason: "Missing collection"
    });
    skipProduct = true;
    continue;
  }

  const collectionHierarchy = resolveIbsherCollection(product.collectionName, browserCatalog);
  if (
    !collectionHierarchy
    || collectionHierarchy.ambiguous
    || !collectionHierarchy.category
    || !collectionHierarchy.subCategory
  ) {
    const reason = collectionHierarchy?.ambiguous
      ? "Collection name matches more than one IBSHER category path"
      : "Collection or its category hierarchy was not found by the API";
    failedItems.dropdownNotFound.push({
      itemcode: product.itemcode,
      name: product.en,
      field: "collection",
      value: product.collectionName,
      reason,
    });
    console.log(`\n❌ SKIPPING: ${reason}: "${product.collectionName}"`);
    continue;
  }

  try {
    // ===== OPEN FORM =====
    await waitForAddProductForm();
    
    await clearAndType(itemCodeInputSelector, product.itemcode);
    await clearAndType('input[name="name.en"]', product.en);
    await clearAndType('input[name="name.ar"]', product.ar);
    await clearAndType('input[name="name.ku"]', product.ku);

    await clearAndType('textarea[name="description.en"]', product.descriptionEn);
    await clearAndType('textarea[name="description.ar"]', product.descriptionAr);
    await clearAndType('textarea[name="description.ku"]', product.descriptionKu);

    await page.select('select[name="currency"]', product.currency === "USD" ? "USD" : "IQD");
    await clearAndType('input[name="price"]', product.price);
    await clearAndType('input[name="discountPrice"]', product.discountPrice);

    // ===== CATEGORY -> SUB-CATEGORY -> COLLECTION =====
    const hierarchyResult = await fillCategoryHierarchy(collectionHierarchy);
    if (!hierarchyResult.success) {
      const reason = `Failed to set ${hierarchyResult.field}: ${hierarchyResult.reason || "unknown"}`;
      failedItems.dropdownNotFound.push({
        itemcode: product.itemcode,
        name: product.en,
        field: hierarchyResult.field,
        value: hierarchyResult.value,
        reason,
      });
      dropdownFailed = true;
      skipProduct = true;
      console.log(`❌ ${reason}`);
      await autoCancelForm();
      continue;
    }

    // ===== BRAND (AUTO-SELECT FROM DROPDOWN) =====
    console.log(`\n🏷️ Setting brand with auto-select...`);
    const brandResult = await fillBrand(product.brand);
    if (!brandResult.success) {
      console.log(`❌ Failed to set brand for ${product.itemcode}`);
      
      // Check if dropdown didn't appear
      if (brandResult.reason === "dropdown_not_appeared") {
        console.log(`⚠️ No dropdown appeared for brand: "${product.brand}"`);
        failedItems.dropdownNotFound.push({
          itemcode: product.itemcode,
          name: product.en,
          field: "brand",
          value: product.brand,
          reason: "Brand dropdown not found"
        });
        dropdownFailed = true;
      } else {
        failedItems.otherErrors.push({
          itemcode: product.itemcode,
          name: product.en,
          reason: `Failed to set brand: ${brandResult.reason || 'unknown'}`
        });
        dropdownFailed = true;
      }
      
      skipProduct = true;
      
      // ===== AUTO CANCEL THE FORM =====
      console.log(`\n🔄 Auto-cancelling form for ${product.itemcode}...`);
      await autoCancelForm();
      console.log(`✅ Form cancelled for ${product.itemcode}`);
      
      continue;
    }

    await new Promise(r => setTimeout(r, 500));

    // ===== KEYWORD =====
    let keywordToUse = product.keyword;
    if (!keywordToUse || keywordToUse === "") {
      keywordToUse = product.en.split(' ').slice(0, 3).join(' ').toLowerCase();
    }
    await fillKeyword(keywordToUse);

    await new Promise(r => setTimeout(r, 300));

    // ===== WARRANTY (AUTO-SELECT FOR UNIT DROPDOWN) =====
    await fillWarranty(
      product.warrantyDuration,
      product.warrantyUnit,
      product.warrantyDescription
    );

    // ===== COLOR OPTIONS, SIZES, STOCK, AND OPTION-SPECIFIC IMAGES =====
    for (let variantIndex = 0; variantIndex < product.variants.length; variantIndex++) {
      const variant = product.variants[variantIndex];
      await ensureVariantOption(variantIndex);
      await tryType(`input[name="variants.${variantIndex}.color.en"]`, variant.color);

      for (let sizeIndex = 0; sizeIndex < variant.sizes.length; sizeIndex++) {
        const size = variant.sizes[sizeIndex];
        await ensureVariantSize(variantIndex, sizeIndex);
        await tryType(`input[name="variants.${variantIndex}.sizes.${sizeIndex}.size"]`, size.size);
        await tryType(
          `input[name="variants.${variantIndex}.sizes.${sizeIndex}.stockQuantity"]`,
          size.stockQuantity,
        );

        // Set this after stock quantity so the requested status remains the final value.
        if (size.stockStatus === "out_of_stock") {
          console.log(`\nSetting stock status for ${variant.color || `option ${variantIndex + 1}`}...`);
          const stockStatusResult = await fillStockStatus(size.stockStatus, variantIndex, sizeIndex);
          if (!stockStatusResult.success) {
            throw new Error(
              `Failed to select stock status for ${variant.color || `option ${variantIndex + 1}`}: `
              + `${stockStatusResult.reason || "unknown"}`,
            );
          }
          await new Promise(r => setTimeout(r, 300));
        }
      }

      if (variant.images.length > 0) {
        const variantDownloadedImages = await uploadVariantImages(product, variant, variantIndex);
        downloadedImages.push(...variantDownloadedImages);
      }
    }

    console.log("\n" + "=".repeat(60));
    console.log("✅ Form filled successfully!");
    console.log("=".repeat(60));

    // ===== SUBMIT BASED ON MODE =====
    if (submissionMode === 'auto') {
      // Auto submit mode
      const submitSuccess = await autoSubmitForm();
      if (!submitSuccess) {
        console.log("Please manually click the SUBMIT button on the website...");
        await waitForManualSubmit();
      }
      
      // Clean up images after submission
      if (downloadedImages.length > 0) {
        cleanupProductImages(downloadedImages);
      }
      
      // Small delay before next product
      await new Promise(r => setTimeout(r, 2000));
      
      // Check if we need to navigate back to add product page for next product
      const currentUrl = page.url();
      if (!currentUrl.includes('/add-product')) {
        console.log("Navigating back to add product page...");
        await page.goto("https://marketplace.ibsher.com/add-product", {
          waitUntil: "networkidle2",
          timeout: 30000
        });
      }
    } else {
      // Manual submit mode
      console.log("   Please review and click SUBMIT on the website.");
      
      // Wait for manual submit on the website
      await waitForManualSubmit();
      
      // Only clean up images AFTER successful submission
      if (downloadedImages.length > 0) {
        cleanupProductImages(downloadedImages);
      }
      
      // Small delay before next product
      await new Promise(r => setTimeout(r, 2000));
      
      // Check if we need to navigate back to add product page for next product
      const currentUrl = page.url();
      if (!currentUrl.includes('/add-product')) {
        console.log("Navigating back to add product page...");
        await page.goto("https://marketplace.ibsher.com/add-product", {
          waitUntil: "networkidle2",
          timeout: 30000
        });
      }
    }

    // ===== MARK AS PROCESSED =====
    processedItemcodes.add(product.itemcode);
    
  } catch (error) {
    console.error(`❌ Error processing product ${product.itemcode}:`, error.message);
    
    if (!failedItems.otherErrors.some(item => item.itemcode === product.itemcode)) {
      failedItems.otherErrors.push({
        itemcode: product.itemcode,
        name: product.en,
        reason: error.message
      });
    }
    
    // ===== AUTO CANCEL THE FORM ON ERROR =====
    console.log(`\n🔄 Auto-cancelling form due to error...`);
    await autoCancelForm();
    console.log(`✅ Form cancelled`);
    
    // Failed products must not leave downloaded image artifacts behind.
    if (downloadedImages.length > 0) {
      cleanupProductImages(downloadedImages);
    }
    
    continue;
  }
}

}

// ===== FINAL REPORT =====
console.log("\n" + "=".repeat(60));
console.log("🎉 ALL PRODUCTS PROCESSED!");
console.log("=".repeat(60));

// Generate and display the report
console.log("\n" + "=".repeat(60));
console.log("📋 FAILED ITEMS REPORT");
console.log("=".repeat(60));

function formatFailedItem(item) {
  const rowText = item.rowNumber ? `Row ${item.rowNumber} - ` : "";
  const reasonText = item.reason ? ` - ${item.reason}` : "";
  return `${rowText}${item.itemcode}: ${item.name}${reasonText}`;
}

const totalFailed = failedItems.missingDetails.length + failedItems.noBrand.length + failedItems.noCollection.length + 
                   failedItems.noBrandAndCollection.length + failedItems.dropdownNotFound.length +
                   failedItems.duplicateItemcode.length + failedItems.invalidProductInfo.length +
                   failedItems.otherErrors.length;


const processedItems = products
  .filter((product) => processedItemcodes.has(product.itemcode))
  .flatMap((product) => flattenProductRows(product));
const reportItems = buildAutomationReportItems(rows, failedItems);
const finalReport = {
  generatedAt: new Date().toISOString(),
  excelFile: SOURCE_EXCEL_FILE,
  sheetName: SHEET_NAME,
  totalRows: rows.length,
  processed: processedItemcodes.size,
  processedItems,
  reportItems,
  failed: totalFailed,
  imageDownloadFailures: failedItems.imageDownloadFailures.length,
  failedItems,
};

console.log(`${AUTOMATION_REPORT_PREFIX}${JSON.stringify(finalReport)}`);

if (totalFailed === 0) {
  console.log("✅ All items were processed successfully!");
} else {
  console.log(`❌ Total failed items: ${totalFailed}\n`);
  
  if (failedItems.duplicateItemcode.length > 0) {
    console.log("🔄 DUPLICATE ITEM CODES:");
    failedItems.duplicateItemcode.forEach(item => {
      console.log(`   - ${formatFailedItem(item)}`);
    });
    console.log("");
  }
  
  if (failedItems.dropdownNotFound.length > 0) {
    console.log("🔍 DROPDOWN NOT FOUND:");
    failedItems.dropdownNotFound.forEach(item => {
      console.log(`   - ${item.itemcode}: ${item.name} - ${item.field}: "${item.value}"`);
    });
    console.log("");
  }
  
  if (failedItems.noBrandAndCollection.length > 0) {
    console.log("🚫 ITEMS WITH NO BRAND AND NO COLLECTION:");
    failedItems.noBrandAndCollection.forEach(item => {
      console.log(`   - ${formatFailedItem(item)}`);
    });
    console.log("");
  }
  
  if (failedItems.noBrand.length > 0) {
    console.log("🚫 ITEMS WITH NO BRAND:");
    failedItems.noBrand.forEach(item => {
      console.log(`   - ${formatFailedItem(item)}`);
    });
    console.log("");
  }
  
  if (failedItems.noCollection.length > 0) {
    console.log("🚫 ITEMS WITH NO COLLECTION:");
    failedItems.noCollection.forEach(item => {
      console.log(`   - ${formatFailedItem(item)}`);
    });
    console.log("");
  }

  if (failedItems.missingDetails.length > 0) {
    console.log("ITEMS WITH MISSING DETAILS:");
    failedItems.missingDetails.forEach(item => {
      console.log(`   - ${formatFailedItem(item)}`);
    });
    console.log("");
  }

  if (failedItems.invalidProductInfo.length > 0) {
    console.log("ITEMS WITH INVALID PRODUCT INFO:");
    failedItems.invalidProductInfo.forEach(item => {
      console.log(`   - ${formatFailedItem(item)}`);
    });
    console.log("");
  }
  
  if (failedItems.otherErrors.length > 0) {
    console.log("⚠️ ITEMS WITH OTHER ERRORS:");
    failedItems.otherErrors.forEach(item => {
      console.log(`   - ${item.itemcode}: ${item.name} - ${item.reason}`);
    });
    console.log("");
  }
}

console.log("Detailed failure report is available in Product Uploader.");

console.log("=".repeat(60));

if (fs.existsSync(DOWNLOAD_DIR)) {
  fs.rmSync(DOWNLOAD_DIR, { recursive: true, force: true });
  console.log("✅ All temporary product images deleted");
}

if (browser) {
  await new Promise(r => setTimeout(r, 2000));
  await browser.close();
}

