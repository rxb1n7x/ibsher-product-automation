const { app, BrowserWindow, clipboard, dialog, ipcMain, nativeImage, net, safeStorage, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const axios = require("axios");
const FormData = require("form-data");
const { spawn } = require("child_process");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const url = require("url");
const packageJson = require("../package.json");
const { LoginRateLimiter } = require("./lib/auth-security");
const {
  getAllowedAccountSections,
  hasAccountSectionAccess,
} = require("./lib/account-sections");
const {
  extractPostimgImageUrls,
  isPostimgPageUrl,
  resolveExcelImageSources,
} = require("./lib/ai-image-source");
const {
  extractEmbeddedImagesFromFirstWorksheet,
  extractEmbeddedImagesFromWorkbook,
} = require("./lib/xlsx-embedded-images");
const {
  buildExistingLocalizationInstructions,
  buildLocalizedSchemaDescriptions,
  getLocalizedGroupState,
} = require("./lib/ai-product-localization");
const productImageSearch = require("./lib/product-image-search");
const ibsherProductAudit = require("./lib/ibsher-product-audit");
const ibsherDataExport = require("./lib/ibsher-data-export");
const ibsherPageImages = require("./lib/ibsher-page-images");
const ibsherBrandImport = require("./lib/ibsher-brand-import");
const ibsherProductPublishing = require("./lib/ibsher-product-publishing");
const ibsherIntegrationReview = require("./lib/ibsher-integration-review");
const ibsherVendorProductDelete = require("./lib/ibsher-vendor-product-delete");
const automationReportExport = require("./lib/automation-report-export");
const {
  buildIbsherProductCatalog,
  collectIbsherApiListPages,
} = require("./lib/ibsher-catalog");
const ibsherProductBulkUpdate = require("./lib/ibsher-product-bulk-update");
const ibsherItemcodeUpdate = require("./lib/ibsher-itemcode-update");
const ibsherExchangeRates = require("./lib/ibsher-exchange-rates");
const ibsherExchangeAccess = require("./lib/ibsher-exchange-access");
const ibsherBugFinder = require("./lib/ibsher-bug-finder");
const {
  getRequestedVendorValue,
  normalizeVendorAccountInput,
  publicVendorAccount,
} = require("./lib/vendor-accounts");
const { normalizeItemcodeField, sortProductRowsByItemcode } = require("./automation/product-row-sort.cjs");
const {
  findProductRowConflicts,
  flattenProductRows,
  normalizeProductPricing,
} = require("./automation/product-variants.cjs");
const { normalizeDescriptionText } = require("./lib/description-text");
const aiProductPerformance = require("./lib/ai-product-performance");
const aiProductKeywords = require("./lib/ai-product-keywords");
const { prepareAiProductImageBuffer } = require("./lib/ai-product-image-format");
const { removeSolidBackground } = require("./lib/solid-background-removal");
const {
  POSTIMAGES_IMAGE_EXTENSIONS,
  extractPostimagesUploadResult,
  resolvePostimagesLocalPath,
} = require("./lib/postimages-upload");

let mainWindow = null;
let automationProcess = null;
const AUTOMATION_REPORT_PREFIX = "__IBSHER_AUTOMATION_REPORT__";
let isCheckingForUpdate = false;
let isDownloadingUpdate = false;
let availableUpdateInfo = null;
let downloadedUpdateInfo = null;

const APP_VERSION = packageJson.version;
const APP_NAME = packageJson.productName || "Ibsher Product Automation";

const TEMPLATE_COLUMNS = [
  "itemCode",
  "en",
  "ar",
  "ku",
  "description-en",
  "description-ar",
  "description-ku",
  "collectionName",
  "brand",
  "price USD",
  "price IQD",
  "discountPrice",
  "keyword",
  "warrantyDuration",
  "warrantyUnit",
  "warrantyDescription",
  "color",
  "size",
  "stockQuantity",
  "image-1",
  "image-2",
  "image-3",
  "image-4",
];

const INTERNAL_PRODUCT_COLUMNS = [
  "itemcode",
  ...TEMPLATE_COLUMNS.filter((column) => column !== "itemCode"),
  "image-5",
  "image-6",
  "image-7",
  "image-8",
  "image-9",
  "image-10",
  "image-11",
  "image-12",
  "image-13",
  "image-14",
  "image-15",
];

const REQUIRED_PRODUCT_COLUMNS = [
  "itemcode",
  "en",
  "ar",
  "ku",
  "collectionName",
  "brand",
  "description-en",
  "description-ar",
  "description-ku",
  "stockQuantity",
];

autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = false;

function getUpdaterLogPath() {
  return path.join(app.getPath("userData"), "logs", "updater.log");
}

function logUpdateEvent(eventName, details = {}) {
  const entry = {
    time: new Date().toISOString(),
    event: eventName,
    details,
  };

  console.log("[updater]", eventName, details);

  try {
    const logPath = getUpdaterLogPath();
    fs.mkdirSync(path.dirname(logPath), { recursive: true });
    fs.appendFileSync(logPath, `${JSON.stringify(entry)}\n`, "utf8");
  } catch (error) {
    console.warn("[updater] Could not write updater log:", error.message);
  }
}

function emitUpdaterStatus(type, payload = {}) {
  const status = {
    type,
    currentVersion: APP_VERSION,
    checking: isCheckingForUpdate,
    downloading: isDownloadingUpdate,
    availableUpdate: availableUpdateInfo,
    downloadedUpdate: downloadedUpdateInfo,
    ...payload,
  };

  logUpdateEvent(type, status);
  sendToRenderer("updater:event", status);
  return status;
}

function sanitizeUpdateInfo(info = {}) {
  return {
    version: info.version,
    releaseDate: info.releaseDate,
    releaseName: info.releaseName,
  };
}

let authSession = null;
let cachedEnv = null;
const loginRateLimiter = new LoginRateLimiter();

function readBundledNeonConfig() {
  const configPaths = [
    path.join(__dirname, "neon-config.json"),
    path.join(process.resourcesPath || "", "neon-config.json"),
  ];

  for (const configPath of configPaths) {
    try {
      if (fs.existsSync(configPath)) {
        return JSON.parse(fs.readFileSync(configPath, "utf8"));
      }
    } catch (error) {
      console.warn(`[config] Could not read ${configPath}: ${error.message}`);
    }
  }

  return {};
}

function readEnvFile() {
  if (cachedEnv) return cachedEnv;

  const envPaths = [
    path.join(process.cwd(), ".env"),
    path.join(__dirname, "..", ".env"),
    path.join(__dirname, ".env"),
    path.join(process.cwd(), ".env.local"),
    path.join(__dirname, "..", ".env.local"),
    path.join(__dirname, ".env.local"),
  ];

  cachedEnv = {};
  for (const envPath of envPaths) {
    if (!fs.existsSync(envPath)) continue;
    const lines = fs.readFileSync(envPath, "utf8").split(/\r?\n/);
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#") || !trimmed.includes("=")) continue;
      const separator = trimmed.indexOf("=");
      const key = trimmed.slice(0, separator).trim();
      const value = trimmed.slice(separator + 1).trim().replace(/^["']|["']$/g, "");
      cachedEnv[key] = value;
    }
  }

  return cachedEnv;
}

function getConfigValue(...keys) {
  const env = readEnvFile();
  const bundledConfig = readBundledNeonConfig();
  for (const key of keys) {
    if (process.env[key]) return process.env[key];
    if (env[key]) return env[key];
    if (bundledConfig[key]) return bundledConfig[key];
  }
  return "";
}

function getNeonConfig() {
  const apiUrl = getConfigValue("NEON_API_URL").replace(/\/+$/, "");
  if (!apiUrl) throw new Error("Neon is not configured. Set NEON_API_URL to the deployed iBsher Neon API URL.");
  const parsed = new URL(apiUrl);
  if (parsed.protocol !== "https:" && !(parsed.protocol === "http:" && ["localhost", "127.0.0.1"].includes(parsed.hostname))) {
    throw new Error("NEON_API_URL must use HTTPS, except for local development.");
  }
  if (parsed.username || parsed.password || parsed.search || parsed.hash || parsed.pathname !== "/") {
    throw new Error("NEON_API_URL must be an API origin without credentials, path or query.");
  }
  return { apiUrl };
}

function getNeonErrorMessage(data) {
  return String(data?.message || data?.error_description || data?.error || "Neon API request failed.");
}

function createNeonRequestError(data, status) {
  const error = new Error(getNeonErrorMessage(data));
  error.status = Number(status || 0);
  error.code = data?.code || data?.error_code || "";
  return error;
}

function encryptSecret(value) {
  const cleanValue = String(value || "");
  if (safeStorage?.isEncryptionAvailable?.()) {
    return {
      mode: "safeStorage",
      value: safeStorage.encryptString(cleanValue).toString("base64"),
    };
  }

  return {
    mode: "base64",
    value: Buffer.from(cleanValue, "utf8").toString("base64"),
  };
}

function decryptSecret(secret) {
  if (!secret) return "";

  try {
    if (typeof secret === "string") return secret;
    if (secret.mode === "safeStorage") {
      return safeStorage.decryptString(Buffer.from(secret.value || "", "base64"));
    }
    if (secret.mode === "base64") {
      return Buffer.from(secret.value || "", "base64").toString("utf8");
    }
  } catch (error) {
    console.warn("[accounts] Could not decrypt stored password:", error.message);
  }

  return "";
}

function getAuthStorePath() {
  return path.join(app.getPath("userData"), "neon-session.json");
}

function getAiSettingsPath() {
  return path.join(app.getPath("userData"), "ai-product-settings.json");
}

function getPostimagesSettingsPath() {
  return path.join(app.getPath("userData"), "postimages-settings.json");
}

function removeLegacyLocalVendorCredentialStores() {
  const legacyPaths = [
    path.join(app.getPath("userData"), "vendor-accounts.json"),
    path.join(app.getPath("userData"), "ibsher-audit-credentials.json"),
  ];
  for (const legacyPath of legacyPaths) {
    try {
      if (fs.existsSync(legacyPath)) fs.unlinkSync(legacyPath);
    } catch (error) {
      console.warn(`[vendor-accounts] Could not remove legacy local credentials: ${error.message}`);
    }
  }
}

const DEFAULT_GEMINI_MODEL = "gemini-3.5-flash-lite";
const FREE_SEARCH_GEMINI_MODEL = DEFAULT_GEMINI_MODEL;
const geminiModelLimitsCache = new Map();

function readAiSettings() {
  try {
    const stored = JSON.parse(fs.readFileSync(getAiSettingsPath(), "utf8"));
    return {
      geminiApiKey: decryptSecret(stored.geminiApiKey),
      geminiModel: DEFAULT_GEMINI_MODEL,
    };
  } catch {
    return {
      geminiApiKey: getConfigValue("GEMINI_API_KEY"),
      geminiModel: DEFAULT_GEMINI_MODEL,
    };
  }
}

function publicAiSettings(settings = readAiSettings()) {
  return {
    geminiConfigured: Boolean(settings.geminiApiKey),
    geminiModel: DEFAULT_GEMINI_MODEL,
  };
}

function saveAiSettings(input = {}) {
  requireAuthenticatedUser();
  const current = readAiSettings();
  const next = {
    geminiApiKey: String(input.geminiApiKey || current.geminiApiKey || "").trim(),
    geminiModel: DEFAULT_GEMINI_MODEL,
  };

  if (!next.geminiApiKey) throw new Error("Google AI Studio Gemini API key is required.");
  fs.writeFileSync(getAiSettingsPath(), JSON.stringify({
    geminiApiKey: encryptSecret(next.geminiApiKey),
    geminiModel: next.geminiModel,
  }, null, 2), "utf8");

  return publicAiSettings(next);
}

function readPostimagesSettings() {
  try {
    const stored = JSON.parse(fs.readFileSync(getPostimagesSettingsPath(), "utf8"));
    return { apiKey: decryptSecret(stored.apiKey), gallery: String(stored.gallery || "").trim() };
  } catch {
    return { apiKey: getConfigValue("POSTIMAGES_API_KEY"), gallery: getConfigValue("POSTIMAGES_GALLERY") };
  }
}

function publicPostimagesSettings(settings = readPostimagesSettings()) {
  return { configured: Boolean(settings.apiKey), gallery: settings.gallery || "" };
}

function savePostimagesSettings(input = {}) {
  requireAccountSectionAccess("postimages-uploader");
  const current = readPostimagesSettings();
  const apiKey = String(input.apiKey || current.apiKey || "").trim();
  const gallery = String(input.gallery ?? current.gallery ?? "").trim();
  if (!apiKey) throw new Error("Postimages API key is required.");
  fs.writeFileSync(getPostimagesSettingsPath(), JSON.stringify({
    apiKey: encryptSecret(apiKey),
    gallery,
  }, null, 2), "utf8");
  return publicPostimagesSettings({ apiKey, gallery });
}

function normalizeGeminiUsage(usageMetadata = {}) {
  const promptTokenCount = Math.max(0, Number(usageMetadata.promptTokenCount || 0));
  const candidatesTokenCount = Math.max(0, Number(usageMetadata.candidatesTokenCount || 0));
  const thoughtsTokenCount = Math.max(0, Number(usageMetadata.thoughtsTokenCount || 0));
  return {
    promptTokenCount,
    candidatesTokenCount,
    thoughtsTokenCount,
    cachedContentTokenCount: Math.max(0, Number(usageMetadata.cachedContentTokenCount || 0)),
    totalTokenCount: Math.max(
      0,
      Number(usageMetadata.totalTokenCount || (promptTokenCount + candidatesTokenCount + thoughtsTokenCount)),
    ),
  };
}

async function getGeminiModelLimits(model, apiKey) {
  const cleanModel = DEFAULT_GEMINI_MODEL;
  if (geminiModelLimitsCache.has(cleanModel)) return geminiModelLimitsCache.get(cleanModel);

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(cleanModel)}`;
  const response = await fetchWithRetry(endpoint, {
    headers: { "x-goog-api-key": apiKey },
  }, `Gemini model ${cleanModel}`, 2);
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data?.error?.message || `Gemini model ${cleanModel} returned HTTP ${response.status}.`);
  }

  const limits = {
    inputTokenLimit: Math.max(0, Number(data.inputTokenLimit || 0)),
    outputTokenLimit: Math.max(0, Number(data.outputTokenLimit || 0)),
  };
  geminiModelLimitsCache.set(cleanModel, limits);
  return limits;
}

async function getAiProductApiStatus() {
  requireAuthenticatedUser();
  const settings = readAiSettings();
  if (!settings.geminiApiKey) {
    return {
      configured: false,
      model: DEFAULT_GEMINI_MODEL,
      limits: null,
    };
  }

  return {
    configured: true,
    model: DEFAULT_GEMINI_MODEL,
    limits: await getGeminiModelLimits(DEFAULT_GEMINI_MODEL, settings.geminiApiKey),
  };
}

function normalizeRole(role, fallback = "admin") {
  const cleanRole = String(role || fallback).toLowerCase();
  return cleanRole === "owner" ? "owner" : "admin";
}

function normalizeAccountStatus(status, fallback = "active") {
  const cleanStatus = String(status || fallback).toLowerCase();
  return ["disable", "disabled", "inactive"].includes(cleanStatus) ? "disabled" : "active";
}

function isMissingVendorColumnError(error) {
  return String(error?.message || "").includes("vendor_account_value");
}

function isMissingAvatarColumnError(error) {
  return String(error?.message || "").includes("avatar_url");
}

function isMissingAllowedSectionsColumnError(error) {
  return String(error?.message || "").includes("allowed_sections");
}

function createVendorColumnMissingError() {
  return new Error("Neon needs the vendor_account_value column. Run neon-schema.sql in Neon SQL Editor.");
}

function createProfileImageMigrationError() {
  return new Error("Neon profile images are not configured. Run neon-schema.sql in Neon SQL Editor.");
}

function createAllowedSectionsMigrationError() {
  return new Error("Neon account sections are not configured. Run neon-schema.sql in Neon SQL Editor.");
}

function publicAccount(record) {
  const role = normalizeRole(record.role);
  return {
    value: record.id,
    label: record.label || record.email,
    email: record.email || "",
    role,
    status: normalizeAccountStatus(record.status),
    vendorAccountValue: record.vendor_account_value || "",
    avatarUrl: record.avatar_url || "",
    allowedSections: getAllowedAccountSections(role, record.allowed_sections),
  };
}

function publicSessionUser(user) {
  if (!user) return null;
  return {
    value: user.value,
    label: user.label,
    email: user.email,
    role: normalizeRole(user.role),
    vendorAccountValue: user.vendorAccountValue || "",
    avatarUrl: user.avatarUrl || "",
    allowedSections: getAllowedAccountSections(user.role, user.allowedSections),
  };
}

async function neonRequest(pathname, options = {}) {
  const { apiUrl } = getNeonConfig();
  const token = options.token === null ? "" : options.token || authSession?.accessToken || "";
  const response = await fetch(`${apiUrl}${pathname}`, {
    method: options.method || "GET",
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...(options.headers || {}),
    },
    signal: AbortSignal.timeout(30000),
    body: options.body === undefined ? undefined : JSON.stringify(options.body),
  });
  const responseText = await response.text();
  let data;
  try { data = responseText ? JSON.parse(responseText) : null; }
  catch { throw new Error("The Neon API returned an invalid response. Check NEON_API_URL."); }
  if (!response.ok) {
    if (response.status === 401 && options.retryOnExpired !== false && options.token !== null && authSession?.refreshToken) {
      await refreshAuthSession();
      return neonRequest(pathname, { ...options, token: authSession.accessToken, retryOnExpired: false });
    }
    throw createNeonRequestError(data, response.status);
  }
  return data;
}

function normalizeAuthData(authData, profile) {
  const expiresIn = Number(authData.expires_in || 3600);
  const user = {
    value: profile.id,
    label: profile.label || profile.email || authData.user?.email,
    email: profile.email || authData.user?.email,
    role: normalizeRole(profile.role),
    vendorAccountValue: profile.vendor_account_value || "",
    avatarUrl: profile.avatar_url || "",
    allowedSections: getAllowedAccountSections(profile.role, profile.allowed_sections),
  };

  return {
    active: true,
    authUserId: authData.user?.id || profile.auth_user_id || "",
    accessToken: authData.access_token,
    refreshToken: authData.refresh_token,
    expiresAt: Date.now() + expiresIn * 1000,
    user,
  };
}

function normalizeRefreshData(refreshData, previousSession) {
  const expiresIn = Number(refreshData.expires_in || 3600);
  return {
    ...previousSession,
    accessToken: refreshData.access_token,
    refreshToken: refreshData.refresh_token || previousSession.refreshToken,
    expiresAt: Date.now() + expiresIn * 1000,
  };
}

function storeAuthSession(session) {
  const payload = {
    authUserId: session.authUserId || "",
    accessToken: session.accessToken,
    refreshToken: session.refreshToken,
    expiresAt: session.expiresAt,
    user: session.user,
  };

  fs.writeFileSync(
    getAuthStorePath(),
    JSON.stringify({ session: encryptSecret(JSON.stringify(payload)) }, null, 2),
    "utf8"
  );
}

function readStoredAuthSession() {
  try {
    const stored = JSON.parse(fs.readFileSync(getAuthStorePath(), "utf8"));
    const sessionJson = decryptSecret(stored.session);
    return JSON.parse(sessionJson);
  } catch {
    return null;
  }
}

function clearStoredAuthSession() {
  try {
    fs.rmSync(getAuthStorePath(), { force: true });
  } catch {
    // Nothing to remove.
  }
}

function invalidateAuthSession() {
  exchangeRateService?.access?.lock(() => exchangeRateService.whenIdle()).catch(() => {});
  authSession = null;
  ibsherProductCatalogCache.clear();
  clearStoredAuthSession();
}

function isUnrecoverableAuthError(error) {
  return [400, 401, 403].includes(Number(error?.status || 0));
}

function createInvalidSessionError(message = "Session expired. Please log in again.") {
  const error = new Error(message);
  error.status = 401;
  error.code = "INVALID_SESSION";
  return error;
}

async function refreshAuthSession() {
  if (!authSession?.refreshToken) {
    throw createInvalidSessionError();
  }

  const refreshData = await neonRequest("/auth/v1/token?grant_type=refresh_token", {
    method: "POST",
    token: null,
    body: { refresh_token: authSession.refreshToken },
  });

  authSession = normalizeRefreshData(refreshData, authSession);
  storeAuthSession(authSession);
  return authSession;
}

async function ensureAuthSession() {
  if (!authSession) {
    authSession = readStoredAuthSession();
  }

  if (!authSession?.user) {
    throw new Error("Please log in with an app user account before using the app.");
  }

  if (!authSession.accessToken || authSession.expiresAt <= Date.now() + 30000) {
    try {
      await refreshAuthSession();
    } catch (error) {
      if (isUnrecoverableAuthError(error)) {
        invalidateAuthSession();
      }
      const sessionError = new Error(error.message || "Session expired. Please log in again.");
      sessionError.status = error.status;
      sessionError.code = error.code;
      throw sessionError;
    }
  }

  return authSession;
}

const APP_USER_COLUMN_SETS = [
  "id,auth_user_id,label,email,role,status,vendor_account_value,avatar_url,allowed_sections",
  "id,auth_user_id,label,email,role,status,vendor_account_value,allowed_sections",
  "id,auth_user_id,label,email,role,status,avatar_url,allowed_sections",
  "id,auth_user_id,label,email,role,status,allowed_sections",
  "id,auth_user_id,label,email,role,status,vendor_account_value,avatar_url",
  "id,auth_user_id,label,email,role,status,vendor_account_value",
  "id,auth_user_id,label,email,role,status,avatar_url",
  "id,auth_user_id,label,email,role,status",
];

async function fetchCompatibleProfiles(query, token) {
  let lastError;
  for (const columns of APP_USER_COLUMN_SETS) {
    try {
      const separator = query.includes("?") ? "&" : "?";
      return await neonRequest(`${query}${separator}select=${columns}`, { token });
    } catch (error) {
      if (
        !isMissingVendorColumnError(error)
        && !isMissingAvatarColumnError(error)
        && !isMissingAllowedSectionsColumnError(error)
      ) {
        throw error;
      }
      lastError = error;
    }
  }
  throw lastError;
}

async function fetchProfiles(token) {
  return fetchCompatibleProfiles("/rest/v1/app_users?order=created_at.asc", token);
}

async function fetchProfileByEmail(email, token) {
  const encodedEmail = encodeURIComponent(String(email || "").trim().toLowerCase());
  const profiles = await fetchCompatibleProfiles(
    `/rest/v1/app_users?email=eq.${encodedEmail}&limit=1`,
    token
  );
  return profiles[0] || null;
}

async function fetchProfileById(id, token) {
  const encodedId = encodeURIComponent(String(id || "").trim());
  const profiles = await fetchCompatibleProfiles(
    `/rest/v1/app_users?id=eq.${encodedId}&limit=1`,
    token
  );
  return profiles[0] || null;
}

async function signUpAppUser(email, password) {
  const authData = await neonRequest("/auth/v1/signup", {
    method: "POST",
    body: {
      email: String(email || "").trim().toLowerCase(),
      password: String(password || ""),
    },
  });

  if (authData.user?.id) return authData.user;
  if (authData.id) return authData;

  throw new Error("The Neon API did not return the new app login ID.");
}

async function validateAuthSession() {
  const session = await ensureAuthSession();
  let authUser;

  try {
    authUser = await neonRequest("/auth/v1/user", {
      token: session.accessToken,
    });
  } catch (error) {
    if (isUnrecoverableAuthError(error)) {
      invalidateAuthSession();
      throw createInvalidSessionError();
    }
    throw new Error(
      error.message || "Could not validate the login session. No protected data was returned."
    );
  }

  const authUserId = String(authUser?.id || "").trim();
  const authEmail = String(authUser?.email || "").trim().toLowerCase();
  const expectedAuthUserId = String(session.authUserId || "").trim();
  const expectedEmail = String(session.user?.email || "").trim().toLowerCase();

  if (
    !authUserId
    || !authEmail
    || (expectedAuthUserId && expectedAuthUserId !== authUserId)
    || (expectedEmail && expectedEmail !== authEmail)
  ) {
    invalidateAuthSession();
    throw createInvalidSessionError("The login token does not match the current app user.");
  }

  let profile;
  try {
    profile = await fetchProfileByEmail(authEmail, session.accessToken);
  } catch (error) {
    if (isUnrecoverableAuthError(error)) {
      invalidateAuthSession();
      throw createInvalidSessionError();
    }
    throw new Error(
      error.message || "Could not validate app access. No protected data was returned."
    );
  }
  if (
    !profile
    || normalizeAccountStatus(profile.status) !== "active"
    || (profile.auth_user_id && profile.auth_user_id !== authUserId)
  ) {
    invalidateAuthSession();
    throw createInvalidSessionError("This app user is missing, disabled, or no longer authorized.");
  }

  authSession.authUserId = authUserId;
  authSession.user = publicAccount(profile);
  storeAuthSession(authSession);
  return authSession;
}

function requireAuthenticatedUser() {
  if (!authSession?.user) {
    throw new Error("Please log in with an app user account before using the app.");
  }
  return authSession.user;
}

function requireAccountSectionAccess(requiredSections) {
  const user = requireAuthenticatedUser();
  if (!hasAccountSectionAccess(user, requiredSections)) {
    throw new Error("Your account does not have access to this section.");
  }
  return user;
}

async function getAuthSession() {
  try {
    await validateAuthSession();
  } catch (error) {
    return {
      active: false,
      user: null,
      ...(authSession || error?.code === "INVALID_SESSION"
        ? { error: error.message || "Could not validate the login session." }
        : {}),
    };
  }

  return {
    active: true,
    user: publicSessionUser(authSession?.user),
  };
}

async function loginUser(credentials = {}) {
  const email = String(credentials.email || "").trim().toLowerCase();
  const password = String(credentials.password || "");

  if (!email || !password) {
    throw new Error("App user email and password are required.");
  }

  const rateLimit = loginRateLimiter.consume(email);
  if (!rateLimit.allowed) {
    const retryAfterSeconds = Math.max(1, Math.ceil(rateLimit.retryAfterMs / 1000));
    const error = new Error(
      `Too many login attempts. Try again in ${Math.ceil(retryAfterSeconds / 60)} minute(s).`
    );
    error.code = "LOGIN_RATE_LIMITED";
    error.retryAfterSeconds = retryAfterSeconds;
    throw error;
  }

  try {
    const authData = await neonRequest("/auth/v1/token?grant_type=password", {
      method: "POST",
      token: null,
      retryOnExpired: false,
      body: { email, password },
    });
    let profile = await fetchProfileByEmail(email, authData.access_token);
    if (!profile) {
      throw new Error("This app user is not allowed. Ask the Owner to add your email in Account Login App Manager.");
    }
    if (profile.status && profile.status !== "active") {
      throw new Error("This app user is disabled.");
    }

    authSession = normalizeAuthData(authData, profile);
    storeAuthSession(authSession);
    await validateAuthSession();
    loginRateLimiter.reset(email);
    return {
      active: true,
      user: publicSessionUser(authSession.user),
    };
  } catch (error) {
    if (authSession?.user?.email === email) {
      invalidateAuthSession();
    }
    throw error;
  }
}

async function logoutUser() {
  try {
    if (exchangeRateService?.access) await exchangeRateService.access.lock(() => exchangeRateService.whenIdle());
  } catch (error) { console.warn("[exchange-rates] Device release could not be confirmed:", error.message); }
  try {
    if (authSession?.accessToken) await neonRequest("/auth/v1/logout", { method: "POST", retryOnExpired: false });
  } catch (error) {
    console.warn("[auth] Remote logout could not be confirmed:", error.message);
  } finally {
    invalidateAuthSession();
  }
  return { active: false, user: null };
}

function getProfileImageInfo(filePath) {
  const cleanPath = String(filePath || "").trim();
  const mimeType = getImageMimeType(cleanPath);
  if (!["image/jpeg", "image/png", "image/webp"].includes(mimeType)) {
    throw new Error("Choose a JPG, PNG, or WebP profile image.");
  }
  if (!fs.existsSync(cleanPath) || !fs.statSync(cleanPath).isFile()) {
    throw new Error("The selected profile image could not be found.");
  }

  const buffer = fs.readFileSync(cleanPath);
  if (!buffer.length) throw new Error("The selected profile image is empty.");
  if (buffer.length > 5 * 1024 * 1024) {
    throw new Error("Profile images must be 5 MB or smaller.");
  }

  return {
    path: cleanPath,
    name: path.basename(cleanPath),
    size: buffer.length,
    mimeType,
    buffer,
    previewUrl: `data:${mimeType};base64,${buffer.toString("base64")}`,
  };
}

async function selectProfileImage() {
  await ensureAuthSession();
  requireAccountSectionAccess("accounts");
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Choose profile image",
    properties: ["openFile"],
    filters: [
      { name: "Profile Images", extensions: ["jpg", "jpeg", "png", "webp"] },
    ],
  });

  if (result.canceled || !result.filePaths[0]) {
    return { canceled: true, image: null };
  }

  const image = getProfileImageInfo(result.filePaths[0]);
  return {
    canceled: false,
    image: {
      path: image.path,
      name: image.name,
      size: image.size,
      mimeType: image.mimeType,
      previewUrl: image.previewUrl,
    },
  };
}

async function ensureProfileImageSupport(token) {
  try {
    await neonRequest("/rest/v1/app_users?select=avatar_url&limit=1", { token });
  } catch (error) {
    if (isMissingAvatarColumnError(error)) throw createProfileImageMigrationError();
    throw error;
  }
}

async function uploadProfileImage(profileId, filePath) {
  if (!String(profileId || "").trim()) throw new Error("Save the account before uploading its profile image.");
  // Profile avatars live with the Neon account row, so no public storage bucket is needed.
  return getProfileImageInfo(filePath).previewUrl;
}

async function listAccounts() {
  await ensureAuthSession();
  requireAccountSectionAccess("accounts");
  const accounts = await fetchProfiles(authSession.accessToken);
  return accounts.map(publicAccount);
}

function isVendorAccountsMigrationError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("vendor_accounts")
    || message.includes("get_vendor_account_credentials")
    || (message.includes("schema cache") && message.includes("vendor"));
}

function createVendorAccountsMigrationError() {
  return new Error("Neon vendor accounts are not configured. Run neon-schema.sql in Neon SQL Editor.");
}

function requireVendorAccountManager() {
  return requireAccountSectionAccess("vendor-accounts");
}

async function fetchVendorAccountRecords() {
  try {
    return await neonRequest(
      "/rest/v1/vendor_accounts?select=id,label,email,status,created_at,updated_at&order=created_at.asc",
      { token: authSession.accessToken },
    );
  } catch (error) {
    if (isVendorAccountsMigrationError(error)) throw createVendorAccountsMigrationError();
    throw error;
  }
}

async function fetchVendorAccountById(accountValue) {
  const value = String(accountValue || "").trim();
  if (!value) return null;
  try {
    const records = await neonRequest(
      `/rest/v1/vendor_accounts?id=eq.${encodeURIComponent(value)}&select=id,label,email,status&limit=1`,
      { token: authSession.accessToken },
    );
    return records[0] || null;
  } catch (error) {
    if (isVendorAccountsMigrationError(error)) throw createVendorAccountsMigrationError();
    throw error;
  }
}

async function listVendorAccounts() {
  await ensureAuthSession();
  requireAuthenticatedUser();
  return (await fetchVendorAccountRecords()).map(publicVendorAccount);
}

async function saveVendorAccount(account = {}) {
  await ensureAuthSession();
  requireVendorAccountManager();
  const value = String(account.value || "").trim();
  const existing = value ? await fetchVendorAccountById(value) : null;
  const normalized = normalizeVendorAccountInput(account, existing);
  const payload = {
    label: normalized.label,
    email: normalized.email,
    status: normalized.status,
    ...(normalized.password ? { password: normalized.password } : {}),
  };

  try {
    if (value) {
      await neonRequest(`/rest/v1/vendor_accounts?id=eq.${encodeURIComponent(value)}`, {
        method: "PATCH",
        token: authSession.accessToken,
        body: payload,
      });
    } else {
      await neonRequest("/rest/v1/vendor_accounts", {
        method: "POST",
        token: authSession.accessToken,
        body: payload,
      });
    }
  } catch (error) {
    if (isVendorAccountsMigrationError(error)) throw createVendorAccountsMigrationError();
    throw error;
  }
  return listVendorAccounts();
}

async function deleteVendorAccount(accountValue) {
  await ensureAuthSession();
  requireVendorAccountManager();
  const cleanValue = String(accountValue || "").trim();
  if (!await fetchVendorAccountById(cleanValue)) {
    throw new Error("Vendor account was not found.");
  }

  const profiles = await fetchProfiles(authSession.accessToken);
  const assignedProfile = profiles.find((profile) => (
    normalizeRole(profile.role) === "admin"
    && String(profile.vendor_account_value || "") === cleanValue
  ));
  if (assignedProfile) {
    throw new Error(`This vendor account is assigned to ${assignedProfile.label || assignedProfile.email}. Reassign that app user before deleting it.`);
  }

  try {
    await neonRequest(`/rest/v1/vendor_accounts?id=eq.${encodeURIComponent(cleanValue)}`, {
      method: "DELETE",
      token: authSession.accessToken,
    });
  } catch (error) {
    if (isVendorAccountsMigrationError(error)) throw createVendorAccountsMigrationError();
    throw error;
  }
  return listVendorAccounts();
}

async function checkVendorAccount(accountValue) {
  await ensureAuthSession();
  requireVendorAccountManager();
  const account = await resolveIbsherVendorAccount({ vendorAccountValue: accountValue });
  await loginIbsherAdmin(account, true);
  return {
    value: account.value,
    status: "active",
    message: `${account.label} can log in to IBSHER.`,
  };
}

async function saveAccount(account = {}) {
  await ensureAuthSession();
  const operator = requireAccountSectionAccess("accounts");

  const value = String(account.value || "").trim();
  const label = String(account.label || "").trim();
  const email = String(account.email || "").trim();
  const role = normalizeRole(account.role);
  const status = normalizeAccountStatus(account.status);
  const password = String(account.password || "").trim();
  const vendorAccountValue = String(account.vendorAccountValue || "").trim();
  const allowedSections = getAllowedAccountSections(role, account.allowedSections);
  const profileImagePath = String(account.profileImagePath || "").trim();

  if (!label || !email) {
    throw new Error("App user name and app login email are required.");
  }
  if (role === "admin" && !allowedSections.length) {
    throw new Error("Choose at least one section for this admin.");
  }
  if (profileImagePath) {
    getProfileImageInfo(profileImagePath);
    await ensureProfileImageSupport(authSession.accessToken);
  }

  const existingProfile = await fetchProfileByEmail(email, authSession.accessToken);
  if (existingProfile && existingProfile.id !== value) {
    throw new Error("Another account already uses this email.");
  }
  const currentProfile = value ? await fetchProfileById(value, authSession.accessToken) : null;
  if (
    normalizeRole(operator.role) !== "owner"
    && (role === "owner" || normalizeRole(currentProfile?.role) === "owner")
  ) {
    throw new Error("Only the Owner can create or edit an Owner account.");
  }
  if (role === "admin") {
    if (!vendorAccountValue) {
      throw new Error("Choose a vendor account for this admin.");
    }
    if (vendorAccountValue !== "*") {
      const vendorAccount = await fetchVendorAccountById(vendorAccountValue);
      if (!vendorAccount || vendorAccount.status !== "active") {
        throw new Error("The selected vendor account is missing or disabled.");
      }
    }
    if (
      normalizeRole(operator.role) !== "owner"
      && operator.vendorAccountValue !== vendorAccountValue
    ) {
      throw new Error("You cannot assign a vendor account that is not assigned to you.");
    }
  }

  const profile = {
    label,
    email: email.toLowerCase(),
    role,
    status,
    vendor_account_value: role === "admin" ? vendorAccountValue : "",
    allowed_sections: allowedSections,
  };

  let savedProfileId = value;
  if (value) {
    if (!currentProfile) {
      throw new Error("Account was not found.");
    }
    await neonRequest(`/rest/v1/app_users?id=eq.${encodeURIComponent(value)}`, {
      method: "PATCH",
      token: authSession.accessToken,
      body: profile,
    }).catch((error) => {
      if (isMissingVendorColumnError(error)) throw createVendorColumnMissingError();
      if (isMissingAllowedSectionsColumnError(error)) throw createAllowedSectionsMigrationError();
      throw error;
    });
  } else {
    if (!password) {
      throw new Error("App login password is required when creating a new account.");
    }
    const authUser = await signUpAppUser(email, password);
    const insertProfile = {
      ...profile,
      ...(authUser?.id ? { auth_user_id: authUser.id } : {}),
    };
    await neonRequest("/rest/v1/app_users", {
      method: "POST",
      token: authSession.accessToken,
      body: insertProfile,
    }).catch((error) => {
      if (isMissingVendorColumnError(error)) throw createVendorColumnMissingError();
      if (isMissingAllowedSectionsColumnError(error)) throw createAllowedSectionsMigrationError();
      throw error;
    });
    const createdProfile = await fetchProfileByEmail(email, authSession.accessToken);
    savedProfileId = createdProfile?.id || "";
  }

  if (profileImagePath) {
    if (!savedProfileId) {
      throw new Error("Account was saved, but its profile ID could not be read for the image upload.");
    }
    const avatarUrl = await uploadProfileImage(savedProfileId, profileImagePath);
    await neonRequest(`/rest/v1/app_users?id=eq.${encodeURIComponent(savedProfileId)}`, {
      method: "PATCH",
      token: authSession.accessToken,
      body: { avatar_url: avatarUrl },
    }).catch((error) => {
      if (isMissingAvatarColumnError(error)) throw createProfileImageMigrationError();
      throw error;
    });
  }

  return listAccounts();
}

async function deleteAccount(accountValue) {
  await ensureAuthSession();
  const operator = requireAccountSectionAccess("accounts");
  const cleanValue = String(accountValue || "");
  if (operator.value === cleanValue) {
    throw new Error("The signed-in account cannot be deleted.");
  }
  const targetProfile = await fetchProfileById(cleanValue, authSession.accessToken);
  if (normalizeRole(operator.role) !== "owner" && normalizeRole(targetProfile?.role) === "owner") {
    throw new Error("Only the Owner can delete an Owner account.");
  }

  await neonRequest(`/rest/v1/app_users?id=eq.${encodeURIComponent(cleanValue)}`, {
    method: "DELETE",
    token: authSession.accessToken,
  });
  return listAccounts();
}

async function checkAccount(accountValue) {
  await ensureAuthSession();
  requireAccountSectionAccess("accounts");
  const accounts = await fetchProfiles(authSession.accessToken);
  const account = accounts.map(publicAccount).find((record) => record.value === accountValue);
  if (!account) {
    throw new Error("Account was not found.");
  }

  return {
    value: account.value,
    status: account.status,
    message: account.status === "active"
      ? `${account.label} is available as ${account.role}.`
      : `${account.label} is ${account.status}.`,
  };
}


function getAppUrl() {
  if (app.isPackaged) {
    return url.format({
      pathname: path.join(__dirname, "index.html"),
      protocol: "file:",
      slashes: true,
    });
  }

  return "http://127.0.0.1:3000";
}

function sendToRenderer(channel, payload) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send(channel, payload);
  }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 760,
    minWidth: 960,
    minHeight: 640,
    title: `${APP_NAME}`,
    backgroundColor: "#08050e",
    show: false,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  const iconPath = path.join(__dirname, "logo192.png");
  if (fs.existsSync(iconPath)) {
    mainWindow.setIcon(iconPath);
  }

  mainWindow.loadURL(getAppUrl());
  mainWindow.setMenu(null);
  mainWindow.once("ready-to-show", () => mainWindow.show());
  mainWindow.webContents.once("did-finish-load", () => {
    setTimeout(() => {
      checkForUpdates("startup");
    }, 1000);
  });

  // Detached Chromium DevTools can emit repeated internal `console.assert`
  // messages to the parent terminal. Keep development output clean unless
  // DevTools was explicitly requested for a debugging session.
  if (!app.isPackaged && process.env.IBSHER_OPEN_DEVTOOLS === "1") {
    mainWindow.webContents.openDevTools({ mode: "detach" });
  }
}

async function checkForUpdates(source = "manual") {
  if (!app.isPackaged) {
    return emitUpdaterStatus("update-not-available", {
      source,
      message: "Update checks run only in the packaged application.",
    });
  }

  if (isCheckingForUpdate || isDownloadingUpdate) {
    return emitUpdaterStatus("update-check-skipped", {
      source,
      message: "An update check or download is already running.",
    });
  }

  isCheckingForUpdate = true;
  availableUpdateInfo = null;
  downloadedUpdateInfo = null;
  emitUpdaterStatus("checking-for-update", { source });

  try {
    return await autoUpdater.checkForUpdates();
  } catch (error) {
    isCheckingForUpdate = false;
    return emitUpdaterStatus("error", {
      source,
      message: error.message || "Could not check for updates.",
    });
  }
}

async function downloadAvailableUpdate() {
  if (!availableUpdateInfo) {
    throw new Error("No update is available to download.");
  }

  if (isDownloadingUpdate) {
    return emitUpdaterStatus("update-download-skipped", {
      message: "The update download is already running.",
    });
  }

  isDownloadingUpdate = true;
  emitUpdaterStatus("download-started", {
    updateInfo: availableUpdateInfo,
  });

  try {
    return await autoUpdater.downloadUpdate();
  } catch (error) {
    isDownloadingUpdate = false;
    throw error;
  }
}

autoUpdater.on("checking-for-update", () => {
  isCheckingForUpdate = true;
  emitUpdaterStatus("checking-for-update");
});

autoUpdater.on("update-available", (info) => {
  isCheckingForUpdate = false;
  availableUpdateInfo = sanitizeUpdateInfo(info);
  emitUpdaterStatus("update-available", {
    updateInfo: availableUpdateInfo,
  });
});

autoUpdater.on("update-not-available", (info) => {
  isCheckingForUpdate = false;
  availableUpdateInfo = null;
  downloadedUpdateInfo = null;
  emitUpdaterStatus("update-not-available", {
    updateInfo: sanitizeUpdateInfo(info),
  });
});

autoUpdater.on("download-progress", (progress) => {
  isDownloadingUpdate = true;
  emitUpdaterStatus("download-progress", {
    progress: {
      percent: Number(progress.percent || 0),
      transferred: progress.transferred,
      total: progress.total,
      bytesPerSecond: progress.bytesPerSecond,
    },
  });
});

autoUpdater.on("update-downloaded", (info) => {
  isCheckingForUpdate = false;
  isDownloadingUpdate = false;
  downloadedUpdateInfo = sanitizeUpdateInfo(info);
  emitUpdaterStatus("update-downloaded", {
    updateInfo: downloadedUpdateInfo,
  });
});

autoUpdater.on("error", (error) => {
  isCheckingForUpdate = false;
  isDownloadingUpdate = false;
  emitUpdaterStatus("error", {
    message: error.message || "Updater error.",
    stack: error.stack,
  });
});

function getAutomationScriptPath() {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, "app.asar.unpacked", "build", "automation", "app.js");
  }

  return path.join(__dirname, "automation", "app.js");
}

function loadXlsx() {
  return require("xlsx");
}

const AI_PRODUCT_FIELDS = [
  "itemcode",
  "en",
  "ar",
  "ku",
  "description-en",
  "description-ar",
  "description-ku",
  "collectionName",
  "brand",
  "keyword",
  "color",
  "size",
];

const AI_EXCEL_COMPLETION_FIELDS = [
  "en",
  "ar",
  "ku",
  "description-en",
  "description-ar",
  "description-ku",
  "collectionName",
  "brand",
  "keyword",
  "color",
  "size",
];


function normalizeIbsherCatalogKey(value) {
  return cleanCell(value)
    .normalize("NFKC")
    .toLocaleLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function getLocalizedIbsherName(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return {
      en: cleanCell(value.en || value.english),
      ar: cleanCell(value.ar || value.arabic),
      ku: cleanCell(value.ku || value.kurdish),
    };
  }

  return {
    en: cleanCell(value),
    ar: "",
    ku: "",
  };
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

async function fetchIbsherAdminListPages(url, credentials, keys = [], options = {}) {
  const limit = Math.max(1, Math.min(500, Number.parseInt(options.limit, 10) || 100));
  const label = cleanCell(options.label) || cleanCell(url).replace(/^\//, "") || "list";
  return collectIbsherApiListPages({
    keys,
    limit,
    extractList: options.extractList,
    getItemKey: options.getItemKey,
    incompleteMessage: options.incompleteMessage
      || `Could not load the complete IBSHER ${label} list.`,
    tooLargeMessage: options.tooLargeMessage
      || `The IBSHER ${label} list is too large to load completely.`,
    requestPage: async (page, pageLimit) => {
      const response = await requestIbsherAdmin({
        method: "GET",
        url,
        params: { page, limit: pageLimit },
      }, credentials);
      return response.data;
    },
  });
}

function sanitizeIbsherProductCatalog(catalog = {}) {
  const categories = Array.isArray(catalog.categories)
    ? [...new Set(catalog.categories.map((item) => cleanCell(
      typeof item === "string" ? item : item?.name
    )).filter(Boolean))]
    : [];
  const collectionsByName = new Map();

  for (const item of Array.isArray(catalog.collections) ? catalog.collections : []) {
    const name = cleanCell(typeof item === "string" ? item : item?.name);
    const key = normalizeIbsherCatalogKey(name);
    if (!key || collectionsByName.has(key)) continue;
    collectionsByName.set(key, {
      name,
      aliases: Array.isArray(item?.aliases) ? item.aliases.map(cleanCell).filter(Boolean) : [],
      subCategory: cleanCell(item?.subCategory),
      category: cleanCell(item?.category),
    });
  }

  return {
    categories,
    collections: [...collectionsByName.values()],
  };
}

function resolveIbsherCollectionName(value, catalog = {}) {
  const candidate = normalizeIbsherCatalogKey(value);
  if (!candidate) return "";

  for (const collection of catalog.collections || []) {
    const names = [collection.name, ...(collection.aliases || [])];
    if (names.some((name) => normalizeIbsherCatalogKey(name) === candidate)) {
      return cleanCell(collection.name);
    }
  }
  return "";
}

async function resolveIbsherVendorAccount(options = {}) {
  await ensureAuthSession();
  requireAuthenticatedUser();
  const requestedVendorValue = getRequestedVendorValue(options);
  try {
    const records = await neonRequest("/rest/v1/rpc/get_vendor_account_credentials", {
      method: "POST",
      token: authSession.accessToken,
      body: { requested_vendor_value: requestedVendorValue || null },
    });
    const record = records?.[0];
    const account = {
      value: cleanCell(record?.value),
      label: cleanCell(record?.label),
      email: cleanCell(record?.email).toLowerCase(),
      password: String(record?.vendor_password || ""),
    };
    if (!account.value || !account.email || !account.password) {
      throw new Error("The selected vendor account has incomplete IBSHER credentials.");
    }
    return account;
  } catch (error) {
    if (isVendorAccountsMigrationError(error)) throw createVendorAccountsMigrationError();
    throw error;
  }
}

function getIbsherApiErrorMessage(error, fallback) {
  const data = error?.response?.data;
  if (typeof data === "string" && data.trim()) return data.trim();
  const baseMessage = cleanCell(data?.message || data?.error || error?.message || fallback);
  const details = [];
  const collectDetails = (value, pathPrefix = "") => {
    if (details.length >= 8 || value === null || value === undefined) return;
    if (typeof value === "string") {
      const message = cleanCell(value);
      if (message && message.toLowerCase() !== baseMessage.toLowerCase()) {
        details.push(pathPrefix ? `${pathPrefix}: ${message}` : message);
      }
      return;
    }
    if (Array.isArray(value)) {
      value.forEach((entry) => collectDetails(entry, pathPrefix));
      return;
    }
    if (typeof value !== "object") return;
    const field = cleanCell(value.path || value.field || value.param || pathPrefix);
    const message = cleanCell(value.message || value.msg || value.error);
    if (message && message.toLowerCase() !== baseMessage.toLowerCase()) {
      details.push(field ? `${field}: ${message}` : message);
    }
    for (const [key, entry] of Object.entries(value)) {
      if (["path", "field", "param", "message", "msg", "error"].includes(key)) continue;
      collectDetails(entry, field || key);
    }
  };
  collectDetails(data?.errors || data?.details || data?.validationErrors);
  const uniqueDetails = [...new Set(details)];
  return uniqueDetails.length ? `${baseMessage}: ${uniqueDetails.join("; ")}` : baseMessage;
}

const IBSHER_PRODUCT_CATALOG_TTL_MS = 5 * 60 * 1000;
const ibsherProductCatalogCache = new Map();
const ibsherAdminSessions = new Map();
const ibsherAdminLoginPromises = new Map();
let activeIbsherAudit = null;

function getIbsherAdminApiBaseUrl() {
  return cleanCell(getConfigValue("IBSHER_API_BASE_URL"))
    .replace(/\/+$/, "")
    || "https://api.ibsher.com/api/v1/admin";
}

function getIbsherAdminCredentials(options = {}) {
  const source = options?.credentials && typeof options.credentials === "object"
    ? options.credentials
    : options;
  const email = cleanCell(source?.email).toLowerCase();
  const password = String(source?.password || "");
  if (!email || !password) {
    throw new Error("Enter the IBSHER email and password in Catalog Audit before loading products.");
  }
  return { email, password };
}

function getIbsherAdminSession(credentials) {
  const key = crypto.createHash("sha256")
    .update(`${credentials.email}\0${credentials.password}`)
    .digest("hex");
  if (!ibsherAdminSessions.has(key)) {
    if (ibsherAdminSessions.size >= 5) {
      const oldestKey = ibsherAdminSessions.keys().next().value;
      ibsherAdminSessions.delete(oldestKey);
      ibsherAdminLoginPromises.delete(oldestKey);
    }
    ibsherAdminSessions.set(key, { cookies: new Map() });
  }
  return { key, session: ibsherAdminSessions.get(key) };
}

function captureIbsherAdminCookies(session, headers = {}) {
  const values = headers["set-cookie"] || headers["Set-Cookie"] || [];
  for (const value of Array.isArray(values) ? values : [values]) {
    const pair = cleanCell(value).split(";", 1)[0];
    const separator = pair.indexOf("=");
    if (separator <= 0) continue;
    const name = pair.slice(0, separator).trim();
    const cookieValue = pair.slice(separator + 1).trim();
    if (cookieValue) session.cookies.set(name, cookieValue);
    else session.cookies.delete(name);
  }
}

function getIbsherAdminCookieHeader(session) {
  return [...session.cookies.entries()]
    .map(([name, value]) => `${name}=${value}`)
    .join("; ");
}

async function loginIbsherAdmin(credentialsInput, force = false) {
  const credentials = getIbsherAdminCredentials(credentialsInput);
  const { key, session } = getIbsherAdminSession(credentials);
  if (!force && session.cookies.size > 0) return session;
  if (!ibsherAdminLoginPromises.has(key)) {
    const loginPromise = (async () => {
      try {
        session.cookies.clear();
        const response = await axios.post(`${getIbsherAdminApiBaseUrl()}/user/login`, credentials, {
          headers: { Accept: "application/json" },
          timeout: 30000,
        });
        captureIbsherAdminCookies(session, response.headers);
        if (!session.cookies.size) {
          throw new Error("IBSHER admin login succeeded but returned no authentication cookie.");
        }
      } catch (error) {
        throw new Error(getIbsherApiErrorMessage(error, "Could not log in to the IBSHER admin API."));
      }
    })().finally(() => {
      ibsherAdminLoginPromises.delete(key);
    });
    ibsherAdminLoginPromises.set(key, loginPromise);
  }
  await ibsherAdminLoginPromises.get(key);
  return session;
}

async function requestIbsherAdmin(configFactory, credentialsInput, retry = true) {
  const credentials = getIbsherAdminCredentials(credentialsInput);
  const session = await loginIbsherAdmin(credentials);
  const createConfig = typeof configFactory === "function" ? configFactory : () => configFactory;
  const config = createConfig();
  config.baseURL = getIbsherAdminApiBaseUrl();
  config.timeout = config.timeout || 60000;
  config.headers = {
    Accept: "application/json",
    ...(config.headers || {}),
    Cookie: getIbsherAdminCookieHeader(session),
  };
  try {
    const response = await axios.request(config);
    captureIbsherAdminCookies(session, response.headers);
    return response;
  } catch (error) {
    if (retry && error?.response?.status === 401) {
      await loginIbsherAdmin(credentials, true);
      return requestIbsherAdmin(configFactory, credentials, false);
    }
    throw new Error(getIbsherApiErrorMessage(error, "IBSHER admin API request failed."));
  }
}

async function getIbsherProductCatalog(options = {}) {
  requireAuthenticatedUser();
  const requestedVendorValue = getRequestedVendorValue(options);
  const cacheKey = `${authSession?.user?.value || "anonymous"}:${requestedVendorValue || "assigned"}`;
  const cached = ibsherProductCatalogCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) return cached.promise;

  const promise = (async () => {
    const credentials = await resolveIbsherVendorAccount(options);
    try {
      const [categoryResponse, subCategories, collections] = await Promise.all([
        requestIbsherAdmin({ method: "GET", url: "/category" }, credentials),
        fetchIbsherAdminListPages("/subcategory", credentials, [
          "subCategory", "subCategories", "subcategory", "subcategories", "data",
        ], { label: "subcategory" }),
        fetchIbsherAdminListPages("/collection", credentials, [
          "collections", "collection", "data",
        ], { label: "collection" }),
      ]);
      const catalog = buildIbsherProductCatalog(
        categoryResponse.data,
        { subCategory: subCategories },
        { collections }
      );
      if (!catalog.categories.length) {
        throw new Error("The IBSHER category API returned no active categories.");
      }
      if (!catalog.subCategories.length) {
        throw new Error("The IBSHER subcategory API returned no active subcategories.");
      }
      if (!catalog.collections.length) {
        throw new Error("The IBSHER collection API returned no active collections.");
      }
      return catalog;
    } catch (error) {
      throw new Error(getIbsherApiErrorMessage(
        error,
        "Could not load categories, subcategories, and collections from IBSHER."
      ));
    }
  })();

  ibsherProductCatalogCache.set(cacheKey, {
    expiresAt: Date.now() + IBSHER_PRODUCT_CATALOG_TTL_MS,
    promise,
  });
  promise.catch(() => {
    if (ibsherProductCatalogCache.get(cacheKey)?.promise === promise) {
      ibsherProductCatalogCache.delete(cacheKey);
    }
  });
  return promise;
}

function normalizeIbsherPagination(value = {}, fallback = {}) {
  return {
    totalProducts: Math.max(0, Number(value.totalProducts || fallback.totalProducts) || 0),
    currentPage: Math.max(1, Number(value.currentPage || fallback.page) || 1),
    totalPages: Math.max(1, Number(value.totalPages) || 1),
    limit: Math.max(1, Number(value.limit || fallback.limit) || 50),
    hasPrevPage: Boolean(value.hasPrevPage),
    hasNextPage: Boolean(value.hasNextPage),
  };
}

function addIbsherProductDisplayUrls(productInput = {}) {
  return ibsherProductAudit.addImageDisplayUrls(productInput, getIbsherAdminApiBaseUrl());
}

async function fetchIbsherProductApiPage(options = {}) {
  const page = Math.max(1, Number.parseInt(options.page, 10) || 1);
  const limit = Math.max(1, Math.min(100, Number.parseInt(options.limit, 10) || 50));
  const search = cleanCell(options.search).slice(0, 160);
  const response = await requestIbsherAdmin({
    method: "GET",
    url: "/product",
    params: { page, limit, ...(search ? { search } : {}) },
  }, options.credentials);
  if (options.requireComplete && (response.data?.success === false
    || !Array.isArray(response.data?.products) || !response.data?.pagination
    || !Number.isInteger(Number(response.data.pagination.totalPages)) || Number(response.data.pagination.totalPages) < 0
    || (Number(response.data.pagination.totalPages) === 0 && Number(response.data.pagination.totalProducts) !== 0)
    || !Number.isInteger(Number(response.data.pagination.totalProducts)) || Number(response.data.pagination.totalProducts) < 0
    || Number(response.data.pagination.currentPage) !== page)) {
    throw new Error("Could not read the complete vendor product list.");
  }
  return {
    products: (Array.isArray(response.data?.products) ? response.data.products : [])
      .map(addIbsherProductDisplayUrls),
    pagination: normalizeIbsherPagination(response.data?.pagination, { page, limit }),
  };
}

function productMatchesIbsherCollection(product, collectionId) {
  const expected = cleanCell(collectionId);
  if (!expected) return true;
  return cleanCell(product?.collectionName?._id || product?.collectionName?.id) === expected;
}

async function getAllIbsherAdminProducts(options = {}, onProgress = () => {}) {
  const firstPage = await fetchIbsherProductApiPage({
    ...options,
    page: 1,
    limit: 100,
  });
  const totalPages = firstPage.pagination.totalPages;
  const products = [...firstPage.products];
  onProgress(products.length, firstPage.pagination.totalProducts);
  if (totalPages > 1) {
    const pages = Array.from({ length: totalPages - 1 }, (_, index) => index + 2);
    const loadedPages = await mapWithConcurrency(pages, 3, async (page) => {
      if (activeIbsherAudit?.canceled) return [];
      const result = await fetchIbsherProductApiPage({ ...options, page, limit: 100 });
      onProgress(result.products.length, firstPage.pagination.totalProducts);
      return result.products;
    });
    loadedPages.forEach((pageProducts) => products.push(...pageProducts));
  }
  const collectionId = cleanCell(options.collectionId);
  return collectionId
    ? products.filter((product) => productMatchesIbsherCollection(product, collectionId))
    : products;
}

async function fetchIbsherProductPage(options = {}) {
  const collectionId = cleanCell(options.collectionId);
  if (!collectionId) return fetchIbsherProductApiPage(options);

  const page = Math.max(1, Number.parseInt(options.page, 10) || 1);
  const limit = Math.max(1, Math.min(100, Number.parseInt(options.limit, 10) || 50));
  const products = await getAllIbsherAdminProducts(options);
  const totalProducts = products.length;
  const totalPages = Math.max(1, Math.ceil(totalProducts / limit));
  const currentPage = Math.min(page, totalPages);
  const start = (currentPage - 1) * limit;
  return {
    products: products.slice(start, start + limit),
    pagination: {
      totalProducts,
      currentPage,
      totalPages,
      limit,
      hasPrevPage: currentPage > 1,
      hasNextPage: currentPage < totalPages,
    },
  };
}

async function getIbsherAdminProducts(options = {}) {
  requireAccountSectionAccess("ibsher-audit");
  const credentials = await resolveIbsherVendorAccount(options);
  return fetchIbsherProductPage({ ...options, credentials });
}

let exchangeRateService = null;
let exchangeRateUserValue = "";
let exchangeRateTimer = null;

function getExchangeRateService() {
  const user = requireAccountSectionAccess("exchange-rates");
  const userValue = cleanCell(user.value);
  if (!userValue) throw new Error("The app user ID is missing.");
  if (exchangeRateService && exchangeRateUserValue === userValue) return exchangeRateService;
  if (exchangeRateService?.running) throw new Error("Wait for the running exchange conversion to finish before switching users.");
  exchangeRateService?.access?.lock().catch(() => {});
  if (exchangeRateTimer) clearInterval(exchangeRateTimer);
  const userKey = crypto.createHash("sha256").update(userValue).digest("hex");
  const statePath = path.join(app.getPath("userData"), "exchange-rates", `${userKey}.json`);
  let state = {};
  try {
    state = JSON.parse(fs.readFileSync(statePath, "utf8"));
    if (state.version !== 1 || !state.vendors || typeof state.vendors !== "object") {
      throw new Error("Invalid exchange-rate storage format.");
    }
  } catch (error) {
    if (error.code !== "ENOENT") throw new Error(`Could not read the saved USD bases. Restore the exchange-rate backup before converting prices: ${error.message}`);
  }
  const authorizeAccount = async () => {
    try {
      await validateAuthSession();
      const currentUser = requireAccountSectionAccess("exchange-rates");
      if (cleanCell(currentUser.value) !== userValue) throw new Error("The exchange-rate user session changed.");
    } catch (error) { error.exchangeFatal = true; throw error; }
  };
  const accessToken = authSession?.accessToken;
  const access = new ibsherExchangeAccess.ExchangeRateAccess({
    canSetup: user.role === "owner", setInterval, clearInterval,
    request: (body) => neonRequest("/rest/v1/rpc/exchange_access", {
      method: "POST", body, token: accessToken, retryOnExpired: false,
    }),
    onState: () => exchangeRateService?.emit(),
  });
  const authorize = async () => { await authorizeAccount(); access.assertUnlocked(); };
  const resolve = async (vendorAccountValue) => {
    await authorize();
    return resolveIbsherVendorAccount({ vendorAccountValue });
  };
  const readProducts = async (credentials, search) => {
    const products = new Map();
    let totalPages;
    let totalProducts;
    for (let page = 1; page <= (totalPages ?? 1); page += 1) {
      await authorize();
      const result = await fetchIbsherProductApiPage({ page, limit: 100, credentials, search, requireComplete: true });
      const pagination = result.pagination;
      if (!Array.isArray(result.products) || pagination?.currentPage !== page
        || !Number.isInteger(pagination.totalPages) || pagination.totalPages < page || pagination.totalPages > 10000
        || !Number.isInteger(pagination.totalProducts) || pagination.totalProducts < 0
        || (totalPages != null && (totalPages !== pagination.totalPages || totalProducts !== pagination.totalProducts))) {
        throw new Error("The vendor product list changed or was incomplete. It will be retried.");
      }
      totalPages = pagination.totalPages;
      totalProducts = pagination.totalProducts;
      for (const product of result.products) {
        const productId = cleanCell(product?._id);
        if (!productId || products.has(productId)) throw new Error("Could not read the complete vendor product list.");
        products.set(productId, product);
      }
    }
    if (products.size !== totalProducts) throw new Error("Could not read the complete vendor product list.");
    return [...products.values()];
  };
  let timerIntervalMs;
  const configureTimer = (status) => {
    if (!status.access?.unlocked) {
      if (exchangeRateTimer) clearInterval(exchangeRateTimer);
      exchangeRateTimer = null;
      timerIntervalMs = undefined;
      return;
    }
    const intervalMs = status.pollIntervalSeconds * 1000;
    if (intervalMs === timerIntervalMs) return;
    if (exchangeRateTimer) clearInterval(exchangeRateTimer);
    timerIntervalMs = intervalMs;
    exchangeRateTimer = setInterval(() => {
      if (cleanCell(authSession?.user?.value) !== userValue || !access.status().unlocked || !exchangeRateService.state.enabled || exchangeRateService.running) return;
      exchangeRateService.check().catch((error) => console.warn(`[exchange-rates] ${error.message}`));
    }, intervalMs);
    exchangeRateTimer.unref?.();
  };
  exchangeRateUserValue = userValue;
  let sharedVendorIds = [];
  let sharedLeaseToken = null;
  let sharedSnapshots = new Map();
  const sharedRequest = async (body) => {
    if (body.action === "release") await authorizeAccount();
    else await authorize();
    return neonRequest("/rest/v1/rpc/exchange_usd_bases", { method: "POST", body: { ...body, runnerToken: access.runnerToken } });
  };
  exchangeRateService = new ibsherExchangeRates.ExchangeRateSync({
    authorize,
    accessStatus: () => access.status(),
    createToken: () => crypto.randomUUID(),
    beginShared: async (vendorIds, { readOnly }) => {
      await access.verify();
      if (!vendorIds.length) return {};
      const leaseToken = readOnly ? null : crypto.randomUUID();
      const result = await sharedRequest({ action: readOnly ? "load" : "acquire", vendorIds, leaseToken });
      if (!readOnly) {
        sharedVendorIds = vendorIds;
        sharedLeaseToken = leaseToken;
        sharedSnapshots = new Map(Object.entries(result.vendors).flatMap(([id, products]) =>
          Object.values(products).map((base) => [`${id}:${base.productId}`, JSON.stringify(base)])));
      }
      return result.vendors;
    },
    endShared: async () => {
      const vendorIds = sharedVendorIds;
      const leaseToken = sharedLeaseToken;
      sharedVendorIds = [];
      sharedLeaseToken = null;
      if (leaseToken) {
        try { await sharedRequest({ action: "release", vendorIds, leaseToken }); }
        catch (error) { console.warn(`[exchange-rates] Could not release Neon lock: ${error.message}`); }
      }
    },
    persist: async (savedState) => {
      // Neon must confirm the USD sources and attempted outputs before any IBSHER write.
      if (sharedLeaseToken) {
        const vendors = Object.fromEntries(sharedVendorIds.map((id) => [id,
          Object.values(savedState.vendors[id]?.products || {}).filter((base) =>
            sharedSnapshots.get(`${id}:${base.productId}`) !== JSON.stringify(base))]));
        await sharedRequest({ action: "save", vendorIds: sharedVendorIds, leaseToken: sharedLeaseToken,
          vendors });
        for (const [id, products] of Object.entries(vendors)) {
          for (const base of products) sharedSnapshots.set(`${id}:${base.productId}`, JSON.stringify(base));
        }
      }
      fs.mkdirSync(path.dirname(statePath), { recursive: true });
      const temporaryPath = `${statePath}.tmp`;
      fs.writeFileSync(temporaryPath, JSON.stringify(savedState, null, 2), "utf8");
      fs.renameSync(temporaryPath, statePath);
    },
    fetchRate: async () => {
      const response = await axios.get(ibsherExchangeRates.SOURCE_URL, {
        timeout: 30000, responseType: "text", maxContentLength: 2 * 1024 * 1024,
        headers: { Accept: "text/html", "Cache-Control": "no-cache" },
        params: { _ibsher: Date.now() },
      });
      return ibsherExchangeRates.parseQamarSellRate(response.data);
    },
    preparePublisher: async () => {
      await authorize();
      const accounts = await listVendorAccounts();
      const matching = accounts.filter((account) => cleanCell(account.label).toLowerCase() === "karim" && account.status === "active");
      if (matching.length !== 1) throw new Error("Exactly one active karim vendor account is required to publish converted products.");
      const credentials = await resolve(matching[0].value);
      return async (product) => {
        await authorize();
        const verifyPrices = (latest) => {
          if (!latest || cleanCell(latest._id) !== cleanCell(product._id) || cleanCell(latest.itemCode) !== cleanCell(product.itemCode)
            || Number(latest.price) !== Number(product.price) || Number(latest.discountPrice || 0) !== Number(product.discountPrice || 0)
            || cleanCell(latest.currency || "IQD").toUpperCase() !== "IQD") {
            throw new Error("The converted product prices could not be verified in the karim account. Publishing remains pending.");
          }
        };
        const latest = await loadIbsherPublishingProduct(product, credentials);
        verifyPrices(latest);
        if (latest.status === "published") return latest;
        if (latest.status !== "review") throw new Error("The converted product is not in Review. Publishing remains pending.");
        await authorize();
        await access.verify();
        const response = await requestIbsherAdmin(ibsherProductPublishing.createProductStatusRequest(product._id, "published"), credentials);
        if (response.data?.success === false) throw new Error(cleanCell(response.data.message) || "IBSHER rejected publication of the converted product.");
        const verified = await loadIbsherPublishingProduct(product, credentials);
        verifyPrices(verified);
        if (verified.status !== "published") throw new Error("IBSHER did not confirm publication of the converted product. Publishing remains pending.");
        return verified;
      };
    },
    listProducts: async (vendorAccountValue) => {
      const credentials = await resolve(vendorAccountValue);
      // A complete snapshot is required before reconciling deleted products.
      const products = await readProducts(credentials);
      return { label: credentials.label, products };
    },
    loadProduct: async (vendorAccountValue, product) => {
      const credentials = await resolve(vendorAccountValue);
      const products = await readProducts(credentials, product.itemCode);
      const exact = products.find((entry) => cleanCell(entry._id) === cleanCell(product._id));
      if (exact) return exact;
      throw new Error("The exact product was not found before price conversion.");
    },
    updateProduct: async (vendorAccountValue, product, target) => {
      const credentials = await resolve(vendorAccountValue);
      await access.verify();
      const candidate = { ...product, ...target, currency: "IQD" };
      // Reuse the established multipart update to retain variants, stock, images and status.
      await requestIbsherAdmin(() => {
        const { form } = createIbsherProductUpdateForm(candidate, { skipRewardValidation: true });
        return { method: "PUT", url: `/product/${encodeURIComponent(product._id)}`, data: form,
          headers: form.getHeaders(), maxContentLength: Infinity, maxBodyLength: Infinity };
      }, credentials);
      return exchangeRateService.deps.loadProduct(vendorAccountValue, product);
    },
    onState: (status) => {
      configureTimer(status);
      if (cleanCell(authSession?.user?.value) === userValue) sendToRenderer("ibsher-exchange:status", status);
    },
  }, state);
  exchangeRateService.access = access;
  configureTimer(exchangeRateService.status());
  return exchangeRateService;
}

async function getIbsherAdminCollections(options = {}) {
  requireAccountSectionAccess("ibsher-audit");
  const credentials = await resolveIbsherVendorAccount(options);
  const collections = await fetchIbsherAdminListPages(
    "/collection",
    credentials,
    ["collections", "collection", "data"],
    { label: "collection" },
  );
  return collections
    .filter((collection) => collection && collection.isActive !== false)
    .map((collection) => {
      const name = getLocalizedIbsherName(collection.name || collection.label);
      const label = name.en || name.ku || name.ar;
      return {
        value: cleanCell(collection._id || collection.id),
        label,
        name,
      };
    })
    .filter((collection) => collection.value && collection.label)
    .sort((left, right) => left.label.localeCompare(right.label));
}

function toIbsherExportFilterOptions(payload, keys = []) {
  return getIbsherApiList(payload, keys)
    .filter((item) => item && item.isActive !== false)
    .map((item) => {
      const name = getLocalizedIbsherName(item.name || item.label);
      return {
        value: cleanCell(item._id || item.id),
        label: name.en || name.ku || name.ar,
        name,
        aliases: [name.en, name.ar, name.ku].filter(Boolean),
      };
    })
    .filter((item) => item.value && item.label)
    .sort((left, right) => left.label.localeCompare(right.label));
}

async function getIbsherExportFilters(options = {}) {
  requireAccountSectionAccess("ibsher-export");
  const credentials = await resolveIbsherVendorAccount(options);
  const [catalog, brands] = await Promise.all([
    getIbsherProductCatalog(options),
    fetchIbsherAdminListPages("/brand", credentials, ["brands", "brand", "data"], { label: "brand" }),
  ]);
  const toOption = (item) => ({
    value: item.id,
    label: item.name,
    name: item.name,
    aliases: item.aliases || [],
    ...(item.categoryId ? { categoryId: item.categoryId } : {}),
    ...(item.subCategoryId ? { subCategoryId: item.subCategoryId } : {}),
  });
  return {
    categories: catalog.categories.map(toOption),
    subCategories: catalog.subCategories.map(toOption),
    collections: catalog.collections.map(toOption),
    brands: toIbsherExportFilterOptions(brands, ["brands", "brand", "data"]),
  };
}

async function exportIbsherProducts(options = {}) {
  requireAccountSectionAccess("ibsher-export");
  const credentials = await resolveIbsherVendorAccount(options);
  const [products, brands] = await Promise.all([
    getAllIbsherAdminProducts({ credentials }),
    // Include inactive brands: existing products can still reference them.
    fetchIbsherAdminListPages("/brand", credentials, ["brands", "brand", "data"], { label: "brand" }),
  ]);
  const exportData = ibsherDataExport.buildIbsherExportRows(products, {
    categoryId: options.categoryId,
    brandId: options.brandId,
    stock: options.stock,
  }, brands);
  if (!exportData.rows.length) {
    throw new Error("No IBSHER products matched the selected export filters.");
  }

  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Export IBSHER products to Excel",
    defaultPath: `ibsher-products-${new Date().toISOString().slice(0, 10)}.xlsx`,
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (result.canceled || !result.filePath) return { canceled: true };

  const xlsx = loadXlsx();
  const workbook = ibsherDataExport.createIbsherExportWorkbook(xlsx, exportData);
  xlsx.writeFile(workbook, result.filePath, { cellStyles: true });

  return {
    canceled: false,
    filePath: result.filePath,
    exportedProducts: exportData.productCount,
    exportedRows: exportData.rows.length,
    imageColumns: exportData.columns.filter((column) => column.startsWith("image-")).length,
  };
}

async function selectIbsherDuplicateExcelFile() {
  requireAccountSectionAccess("ibsher-export");
  return dialog.showOpenDialog(mainWindow, {
    title: "Choose an Excel workbook",
    properties: ["openFile"],
    filters: [
      { name: "Excel Workbooks", extensions: ["xlsx", "xls", "csv"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
}

function parseIbsherItemcodeUpdateWorkbook(filePath) {
  const resolvedPath = requireSpreadsheetFile(filePath, "The item code update");
  const xlsx = loadXlsx();
  const workbook = xlsx.readFile(resolvedPath, { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("The item code update Excel workbook has no worksheet.");
  const matrix = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: false,
  });
  return {
    filePath: resolvedPath,
    fileName: path.basename(resolvedPath),
    sheetName,
    ...ibsherItemcodeUpdate.parseItemcodeUpdateMatrix(matrix),
  };
}

async function selectIbsherItemcodeUpdateFile() {
  requireAccountSectionAccess("ibsher-export");
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Choose the IBSHER item code update Excel workbook",
    properties: ["openFile"],
    filters: [
      { name: "Excel Workbooks", extensions: ["xlsx", "xls", "csv"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
  if (result.canceled || !result.filePaths?.[0]) return { canceled: true };
  const parsed = parseIbsherItemcodeUpdateWorkbook(result.filePaths[0]);
  const { rows, ...preview } = parsed;
  return { canceled: false, ...preview };
}

async function updateIbsherItemcodes(options = {}) {
  requireAccountSectionAccess("ibsher-export");
  if (activeIbsherAudit) throw new Error("Wait for the running IBSHER product audit to finish first.");

  let rows;
  if (cleanCell(options.filePath)) {
    const parsed = parseIbsherItemcodeUpdateWorkbook(options.filePath);
    if (parsed.validationErrors.length) {
      throw new Error(`The item code update workbook is not valid: ${parsed.validationErrors.join(" ")}`);
    }
    rows = parsed.rows;
  } else {
    const oldItemcode = cleanCell(options.oldItemcode);
    const newItemcode = cleanCell(options.newItemcode);
    if (!oldItemcode || !newItemcode) throw new Error("Enter both the current and new item code.");
    if (oldItemcode.toLowerCase() === newItemcode.toLowerCase()) {
      throw new Error("The new item code must be different from the current item code.");
    }
    rows = [{ rowNumber: 1, oldItemcode, newItemcode }];
  }
  if (rows.length > 10000) throw new Error("One item code update is limited to 10,000 products.");

  const credentials = await resolveIbsherVendorAccount(options);
  const results = [];
  let updated = 0;
  let failed = 0;

  await mapWithConcurrency(rows, 2, async (row) => {
    try {
      const oldPage = await fetchIbsherProductPage({ search: row.oldItemcode, limit: 100, credentials });
      const matches = oldPage.products.filter((product) => (
        cleanCell(product.itemCode).toLowerCase() === row.oldItemcode.toLowerCase()
      ));
      if (!matches.length) throw new Error("No product with this exact current item code was found in IBSHER.");
      if (matches.length > 1) throw new Error("More than one product has this current item code; update it manually in IBSHER.");

      const newPage = await fetchIbsherProductPage({ search: row.newItemcode, limit: 100, credentials });
      const conflict = newPage.products.find((product) => (
        product._id !== matches[0]._id
        && cleanCell(product.itemCode).toLowerCase() === row.newItemcode.toLowerCase()
      ));
      if (conflict) throw new Error("The new item code is already used by another IBSHER product.");

      const product = ibsherProductAudit.normalizeIbsherProduct(matches[0]);
      product.itemCode = row.newItemcode;
      const saved = await updateIbsherAdminProductWithCredentials(product, credentials);
      if (cleanCell(saved.product?.itemCode).toLowerCase() !== row.newItemcode.toLowerCase()) {
        throw new Error("IBSHER did not return the updated item code during verification.");
      }
      updated += 1;
      results.push({ ...row, status: "updated" });
    } catch (error) {
      failed += 1;
      results.push({
        ...row,
        status: "failed",
        reason: getIbsherApiErrorMessage(error, "Item code update failed."),
      });
    }
  });

  return {
    total: rows.length,
    updated,
    failed,
    failures: results.filter((result) => result.status === "failed"),
  };
}

function parseIbsherBulkUpdateWorkbook(filePath) {
  const resolvedPath = requireSpreadsheetFile(filePath, "The product update");
  const xlsx = loadXlsx();
  const workbook = xlsx.readFile(resolvedPath, { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("The product update Excel workbook has no worksheet.");
  const matrix = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: true,
  });
  return {
    filePath: resolvedPath,
    fileName: path.basename(resolvedPath),
    sheetName,
    matrix,
    ...ibsherProductBulkUpdate.parseBulkUpdateMatrix(matrix),
  };
}

async function selectIbsherBulkUpdateFile() {
  requireAccountSectionAccess("ibsher-export");
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Choose the IBSHER product update Excel workbook",
    properties: ["openFile"],
    filters: [
      { name: "Excel Workbooks", extensions: ["xlsx", "xls", "csv"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
  if (result.canceled || !result.filePaths?.[0]) return { canceled: true };
  const parsed = parseIbsherBulkUpdateWorkbook(result.filePaths[0]);
  const { rows, matrix, ...preview } = parsed;
  return { canceled: false, ...preview };
}

const BULK_UPDATE_IMAGE_EXTENSIONS = new Set([
  ".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".avif", ".heic", ".heif",
]);
const BULK_UPDATE_MAX_IMAGES_PER_VARIANT = 15;
const BULK_UPDATE_MAX_IMAGE_BYTES_PER_VARIANT = 18 * 1024 * 1024;

function bulkUpdateImageExtension(source, contentType) {
  const normalizedType = cleanCell(contentType).split(";")[0].toLowerCase();
  const byType = {
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/avif": ".avif",
    "image/heic": ".heic",
    "image/heif": ".heif",
  }[normalizedType];
  if (byType) return byType;
  try {
    const extension = path.extname(new URL(source).pathname).toLowerCase();
    return BULK_UPDATE_IMAGE_EXTENSIONS.has(extension) ? extension : ".jpg";
  } catch {
    return ".jpg";
  }
}

async function downloadBulkUpdateImage(source, outputDirectory, filePrefix, imageIndex) {
  const response = await fetchWithRetry(source, {
    headers: {
      Accept: "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
      "User-Agent": PRODUCT_SEARCH_HEADERS["User-Agent"],
    },
    redirect: "follow",
  }, "Product image download", 3);
  if (!response.ok) throw new Error(`Image download returned HTTP ${response.status}: ${source}`);
  const contentType = cleanCell(response.headers.get("content-type")).toLowerCase();
  if (!contentType.startsWith("image/")) {
    if (isPostimgPageUrl(source) && contentType.includes("text/html")) {
      const imageUrls = extractPostimgImageUrls(await response.text(), response.url || source, BULK_UPDATE_MAX_IMAGES_PER_VARIANT);
      if (!imageUrls.length) throw new Error(`No downloadable images were found at ${source}.`);
      const paths = [];
      for (let index = 0; index < imageUrls.length; index += 1) {
        paths.push(...await downloadBulkUpdateImage(
          imageUrls[index],
          outputDirectory,
          filePrefix,
          imageIndex + index,
        ));
      }
      return paths;
    }
    throw new Error(`The image source did not return an image (${contentType || "unknown content type"}): ${source}`);
  }
  const buffer = Buffer.from(await response.arrayBuffer());
  if (!buffer.length) throw new Error(`The downloaded image is empty: ${source}`);
  const extension = bulkUpdateImageExtension(response.url || source, contentType);
  const fingerprint = crypto.createHash("sha256").update(buffer).digest("hex").slice(0, 16);
  const outputPath = path.join(outputDirectory, `${filePrefix}-${imageIndex + 1}-${fingerprint}${extension}`);
  fs.writeFileSync(outputPath, buffer);
  return [outputPath];
}

async function prepareBulkUpdateVariantImages(update, options = {}) {
  const resolved = resolveExcelImageSources(update.sources, {
    workbookFile: options.workbookFile,
    imageExtensions: BULK_UPDATE_IMAGE_EXTENSIONS,
  });
  if (resolved.errors.length) throw new Error(resolved.errors.join(" "));
  if (!resolved.sources.length) throw new Error(`Variant ${update.variantIndex + 1} has no readable replacement images.`);

  const imagePaths = [];
  for (let sourceIndex = 0; sourceIndex < resolved.sources.length; sourceIndex += 1) {
    const source = resolved.sources[sourceIndex];
    if (/^https?:\/\//i.test(source)) {
      imagePaths.push(...await downloadBulkUpdateImage(
        source,
        options.outputDirectory,
        `${cleanCell(options.itemcode).replace(/[^a-z0-9_-]+/gi, "_").slice(0, 60) || "product"}-v${update.variantIndex + 1}`,
        imagePaths.length,
      ));
    } else {
      imagePaths.push(source);
    }
    if (imagePaths.length > BULK_UPDATE_MAX_IMAGES_PER_VARIANT) {
      throw new Error(`Variant ${update.variantIndex + 1} contains more than ${BULK_UPDATE_MAX_IMAGES_PER_VARIANT} replacement images.`);
    }
  }

  const totalBytes = imagePaths.reduce((sum, imagePath) => {
    if (!fs.existsSync(imagePath) || !fs.statSync(imagePath).isFile()) {
      throw new Error(`Replacement image could not be found: ${imagePath}`);
    }
    return sum + fs.statSync(imagePath).size;
  }, 0);
  if (totalBytes > BULK_UPDATE_MAX_IMAGE_BYTES_PER_VARIANT) {
    throw new Error(`Variant ${update.variantIndex + 1} replacement images exceed 18 MB in total.`);
  }
  return imagePaths;
}

async function bulkUpdateIbsherProducts(options = {}) {
  requireAccountSectionAccess("ibsher-export");
  if (activeIbsherAudit) throw new Error("Wait for the running IBSHER product audit to finish first.");
  const parsed = parseIbsherBulkUpdateWorkbook(options.filePath);
  if (parsed.validationErrors.length) {
    throw new Error(`The update workbook is not valid: ${parsed.validationErrors.join(" ")}`);
  }
  if (parsed.rowCount > 50000 || parsed.productCount > 10000) {
    throw new Error("One bulk update is limited to 5,000 rows and 1,000 products.");
  }
  const temporaryImageDirectory = parsed.imageColumns.length
    ? fs.mkdtempSync(path.join(app.getPath("temp"), "ibsher-bulk-update-"))
    : "";

  try {
    if (temporaryImageDirectory && path.extname(parsed.filePath).toLowerCase() === ".xlsx") {
      const imageColumnIndexes = parsed.matrix[0].flatMap((header, columnIndex) => (
        ibsherProductBulkUpdate.IMAGE_COLUMNS.includes(ibsherProductBulkUpdate.canonicalColumn(header))
          ? [columnIndex]
          : []
      ));
      const embeddedImagesByRow = await extractEmbeddedImagesFromFirstWorksheet(
        parsed.filePath,
        temporaryImageDirectory,
        imageColumnIndexes,
        parsed.sheetName,
      );
      parsed.rows.forEach((row) => {
        row.__embeddedImages = embeddedImagesByRow.get(row.__rowNumber) || [];
      });
    }

    const credentials = await resolveIbsherVendorAccount(options);
    const [catalog, brandEntries] = await Promise.all([
      getIbsherProductCatalog(options),
      fetchIbsherAdminListPages("/brand", credentials, ["brands", "brand", "data"], { label: "brand" }),
    ]);
    const brands = toIbsherExportFilterOptions(brandEntries, ["brands", "brand", "data"])
      .map((brand) => ({ id: brand.value, name: brand.name, label: brand.label, aliases: brand.aliases }));
    const groups = ibsherProductBulkUpdate.groupRowsByItemcode(parsed.rows);
    const results = [];
    let processed = 0;
    let updated = 0;
    let failed = 0;
    let skipped = 0;

    await mapWithConcurrency(groups, 2, async (group) => {
      let outcome;
      try {
        const page = await fetchIbsherProductPage({ search: group.itemcode, limit: 50, credentials });
        const product = page.products.find((candidate) => (
          cleanCell(candidate.itemCode).toLowerCase() === cleanCell(group.itemcode).toLowerCase()
        ));
        if (!product) throw new Error("No product with this exact itemcode was found in IBSHER.");
        if (!ibsherProductBulkUpdate.productMatchesScope(product, options)) {
          outcome = { itemcode: group.itemcode, status: "skipped", reason: "Product is outside the selected update scope." };
          skipped += 1;
        } else {
          const normalized = ibsherProductAudit.normalizeIbsherProduct(product);
          const update = ibsherProductBulkUpdate.applyBulkUpdateRows(
            normalized,
            group.rows,
            parsed.updateColumns,
            { catalog, brands },
          );
          const imageUpdates = ibsherProductBulkUpdate.buildVariantImageUpdates(
            normalized,
            group.rows,
            parsed.imageColumns,
          );
          const replacementImagesByVariant = new Map();
          for (const imageUpdate of imageUpdates) {
            replacementImagesByVariant.set(
              imageUpdate.variantIndex,
              await prepareBulkUpdateVariantImages(imageUpdate, {
                workbookFile: parsed.filePath,
                outputDirectory: temporaryImageDirectory,
                itemcode: group.itemcode,
              }),
            );
          }
          const validationErrors = getIbsherProductUpdateValidationErrors(update.product, {
            replacementImagesByVariant,
          });
          if (validationErrors.length) {
            throw new Error(`Product validation failed: ${validationErrors.slice(0, 8).join("; ")}`);
          }
          await updateIbsherAdminProductWithCredentials(update.product, credentials, {
            replacementImagesByVariant,
          });
          const imageChanges = imageUpdates.map(({ variantIndex }) => `images:${variantIndex + 1}`);
          outcome = {
            itemcode: group.itemcode,
            status: "updated",
            changes: [...update.changes, ...imageChanges],
          };
          updated += 1;
        }
      } catch (error) {
        outcome = { itemcode: group.itemcode, status: "failed", reason: getIbsherApiErrorMessage(error, "Product update failed.") };
        failed += 1;
      }
      results.push(outcome);
      processed += 1;
      sendToRenderer("ibsher-bulk-update:progress", {
        vendorAccountValue: getRequestedVendorValue(options),
        processed,
        total: groups.length,
        updated,
        failed,
        skipped,
        itemcode: group.itemcode,
      });
    });

    return {
      total: groups.length,
      updated,
      failed,
      skipped,
      columns: parsed.updateColumns,
      failures: results.filter((result) => result.status !== "updated"),
    };
  } finally {
    if (temporaryImageDirectory) {
      try {
        fs.rmSync(temporaryImageDirectory, {
          recursive: true,
          force: true,
          maxRetries: 3,
          retryDelay: 100,
        });
      } catch (cleanupError) {
        console.warn(`[ibsher-bulk-update] Could not remove temporary images: ${cleanupError.message}`);
      }
    }
  }
}

async function exportIbsherBulkUpdateIssues(options = {}) {
  requireAccountSectionAccess("ibsher-export");
  const sourceFile = requireSpreadsheetFile(options.filePath, "The product update");
  const xlsx = loadXlsx();
  const sourceWorkbook = xlsx.readFile(sourceFile, { cellDates: true });
  const sourceSheetName = sourceWorkbook.SheetNames[0];
  if (!sourceSheetName) throw new Error("The product update Excel workbook has no worksheet.");

  const sourceMatrix = xlsx.utils.sheet_to_json(sourceWorkbook.Sheets[sourceSheetName], {
    header: 1,
    defval: "",
    raw: true,
  });
  const issueExport = ibsherProductBulkUpdate.buildBulkUpdateIssueMatrix(
    sourceMatrix,
    options.failures,
  );
  const sourceName = path.parse(sourceFile).name;
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Export products with update issues",
    defaultPath: path.join(path.dirname(sourceFile), `${sourceName}-issues-${timestamp}.xlsx`),
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (result.canceled || !result.filePath) return { canceled: true };

  const worksheet = xlsx.utils.aoa_to_sheet(issueExport.matrix);
  worksheet["!autofilter"] = { ref: worksheet["!ref"] };
  worksheet["!cols"] = issueExport.matrix[0].map((header, columnIndex) => {
    const width = issueExport.matrix.reduce((maximum, row) => (
      Math.max(maximum, cleanCell(row[columnIndex]).length)
    ), cleanCell(header).length);
    return { wch: Math.min(Math.max(width + 2, 12), 72) };
  });
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Update Issues");
  xlsx.writeFile(workbook, result.filePath, { cellStyles: true });

  return {
    canceled: false,
    filePath: result.filePath,
    exportedProducts: issueExport.productCount,
    exportedRows: issueExport.rowCount,
  };
}

function requireSpreadsheetFile(filePath, label) {
  const resolvedPath = path.resolve(cleanCell(filePath));
  if (!filePath || !fs.existsSync(resolvedPath) || !fs.statSync(resolvedPath).isFile()) {
    throw new Error(`${label} Excel workbook could not be found.`);
  }
  return resolvedPath;
}

async function cleanIbsherDuplicateItemCodes(options = {}) {
  requireAccountSectionAccess("ibsher-export");
  const referenceFile = requireSpreadsheetFile(options.referenceFile, "The first");
  const targetFile = requireSpreadsheetFile(options.targetFile, "The second");
  if (referenceFile.toLowerCase() === targetFile.toLowerCase()) {
    throw new Error("Choose two different Excel workbooks.");
  }

  const xlsx = loadXlsx();
  const readOptions = { cellDates: true, cellNF: true, cellStyles: true };
  const referenceWorkbook = xlsx.readFile(referenceFile, readOptions);
  const targetWorkbook = xlsx.readFile(targetFile, readOptions);
  const cleanup = ibsherDataExport.removeDuplicateItemCodeRows(
    xlsx,
    referenceWorkbook,
    targetWorkbook,
  );

  const targetName = path.parse(targetFile).name;
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Save the Excel workbook without duplicate item codes",
    defaultPath: `${targetName}-without-duplicates.xlsx`,
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (result.canceled || !result.filePath) return { canceled: true };
  const outputFile = path.resolve(result.filePath);
  if ([referenceFile, targetFile].some((filePath) => filePath.toLowerCase() === outputFile.toLowerCase())) {
    throw new Error("Save the cleaned workbook as a new file so the uploaded files stay unchanged.");
  }

  xlsx.writeFile(cleanup.workbook, outputFile, {
    bookType: "xlsx",
    cellStyles: true,
  });

  return {
    canceled: false,
    filePath: outputFile,
    referenceCodeCount: cleanup.referenceCodeCount,
    targetRows: cleanup.targetRows,
    removedRows: cleanup.removedRows,
    matchedCodes: cleanup.matchedCodes,
    remainingRows: cleanup.remainingRows,
    referenceSheet: cleanup.referenceSheet,
    targetSheet: cleanup.targetSheet,
  };
}

function getIbsherAuditStatus() {
  requireAccountSectionAccess("ibsher-audit");
  return {
    ai: publicAiSettings(),
    running: Boolean(activeIbsherAudit),
    runId: activeIbsherAudit?.id || "",
    processed: activeIbsherAudit?.processed || 0,
    total: activeIbsherAudit?.total || 0,
    operation: activeIbsherAudit?.operation || "",
  };
}

function extensionForImageContentType(contentType = "") {
  const mime = cleanCell(contentType).split(";", 1)[0].toLowerCase();
  return {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/gif": ".gif",
    "image/avif": ".avif",
  }[mime] || ".img";
}

async function downloadImageForBackgroundRemoval(image, directory, prefix) {
  const source = cleanCell(image?.displayUrl || image?.url);
  if (!source) throw new Error("The image has no downloadable URL.");
  const response = await axios.get(source, {
    responseType: "arraybuffer",
    timeout: 60000,
    maxContentLength: 25 * 1024 * 1024,
    maxBodyLength: 25 * 1024 * 1024,
    headers: { Accept: "image/avif,image/webp,image/png,image/jpeg,image/*;q=0.8" },
  });
  const contentType = cleanCell(response.headers?.["content-type"]);
  if (!contentType.toLowerCase().startsWith("image/")) {
    throw new Error(`The image URL returned ${contentType || "non-image data"}.`);
  }
  const buffer = Buffer.from(response.data);
  if (!buffer.length) throw new Error("The downloaded image is empty.");
  const originalPath = path.join(directory, `${prefix}${extensionForImageContentType(contentType)}`);
  fs.writeFileSync(originalPath, buffer);
  return { buffer, originalPath };
}

async function removeIbsherSolidBackgrounds(options = {}) {
  requireAccountSectionAccess("ibsher-audit");
  if (activeIbsherAudit) throw new Error("Wait for the running catalog action to finish first.");
  const credentials = await resolveIbsherVendorAccount(options);
  const run = {
    id: `background-removal-${Date.now()}`,
    operation: "background-removal",
    vendorAccountValue: credentials.value,
    canceled: false,
    processed: 0,
    total: 0,
    loaded: 0,
  };
  activeIbsherAudit = run;
  const summary = {
    checkedProducts: 0,
    checkedImages: 0,
    updatedProducts: 0,
    backgroundsRemoved: 0,
    alreadyTransparent: 0,
    complexBackgrounds: 0,
    failedProducts: 0,
  };
  const failures = [];
  const temporaryDirectory = fs.mkdtempSync(path.join(app.getPath("temp"), "ibsher-background-removal-"));
  try {
    emitIbsherAuditProgress(run, {
      operation: run.operation,
      phase: "loading",
      model: "",
      message: "Loading every product from IBSHER...",
      summary,
    });
    const products = await getAllIbsherAdminProducts({ ...options, credentials }, (count, total) => {
      run.total = total;
      run.loaded = Math.min(total, run.loaded + count);
      emitIbsherAuditProgress(run, {
        operation: run.operation,
        phase: "loading",
        loaded: run.loaded,
        model: "",
        message: `Loaded ${run.loaded} of ${total} products...`,
        summary,
      });
    });
    run.total = products.length;

    for (let productIndex = 0; productIndex < products.length; productIndex += 1) {
      if (run.canceled) break;
      const product = products[productIndex];
      const productDirectory = path.join(temporaryDirectory, `product-${productIndex}`);
      fs.mkdirSync(productDirectory, { recursive: true });
      const replacementImagesByVariant = new Map();
      let productRemovedCount = 0;
      let productFailed = false;
      try {
        for (let variantIndex = 0; variantIndex < product.variants.length; variantIndex += 1) {
          const variant = product.variants[variantIndex];
          const preparedPaths = [];
          let variantRemovedCount = 0;
          let variantFailed = false;
          for (let imageIndex = 0; imageIndex < variant.images.length; imageIndex += 1) {
            if (run.canceled) break;
            try {
              const downloaded = await downloadImageForBackgroundRemoval(
                variant.images[imageIndex],
                productDirectory,
                `variant-${variantIndex}-image-${imageIndex}`,
              );
              summary.checkedImages += 1;
              const result = await removeSolidBackground(downloaded.buffer);
              if (result.changed) {
                const outputPath = path.join(productDirectory, `variant-${variantIndex}-image-${imageIndex}-transparent.png`);
                fs.writeFileSync(outputPath, result.buffer);
                preparedPaths.push(outputPath);
                variantRemovedCount += 1;
              } else {
                preparedPaths.push(downloaded.originalPath);
                if (result.reason === "already-transparent") summary.alreadyTransparent += 1;
                else summary.complexBackgrounds += 1;
              }
            } catch (error) {
              variantFailed = true;
              productFailed = true;
              failures.push({
                itemCode: product.itemCode,
                variant: variantIndex + 1,
                image: imageIndex + 1,
                error: error.message,
              });
              break;
            }
          }
          if (!run.canceled && !variantFailed && variantRemovedCount > 0
            && preparedPaths.length === variant.images.length) {
            replacementImagesByVariant.set(variantIndex, preparedPaths);
            productRemovedCount += variantRemovedCount;
          }
        }

        if (!run.canceled && replacementImagesByVariant.size > 0) {
          await updateIbsherAdminProductWithCredentials(product, credentials, { replacementImagesByVariant });
          summary.updatedProducts += 1;
          summary.backgroundsRemoved += productRemovedCount;
        }
      } catch (error) {
        productFailed = true;
        failures.push({ itemCode: product.itemCode, error: error.message });
      } finally {
        if (productFailed) summary.failedProducts += 1;
        summary.checkedProducts += 1;
        run.processed = productIndex + 1;
        fs.rmSync(productDirectory, { recursive: true, force: true });
      }
      emitIbsherAuditProgress(run, {
        operation: run.operation,
        phase: "background-removal",
        model: "",
        summary: { ...summary },
        message: `Checked ${run.processed} of ${run.total} products; removed ${summary.backgroundsRemoved} backgrounds.`,
      });
    }

    const phase = run.canceled ? "canceled" : "complete";
    emitIbsherAuditProgress(run, {
      operation: run.operation,
      phase,
      model: "",
      summary: { ...summary },
      message: run.canceled
        ? `Background removal stopped after ${run.processed} products.`
        : `Background removal completed for ${run.processed} products.`,
    });
    return { canceled: run.canceled, summary, failures: failures.slice(0, 500) };
  } catch (error) {
    emitIbsherAuditProgress(run, {
      operation: run.operation,
      phase: "error",
      model: "",
      message: error.message,
      summary,
    });
    throw error;
  } finally {
    fs.rmSync(temporaryDirectory, { recursive: true, force: true });
    if (activeIbsherAudit?.id === run.id) activeIbsherAudit = null;
  }
}

function getAuditSummary(results = []) {
  return results.reduce((summary, result) => {
    summary.checked += 1;
    if (result.status === "error") summary.errors += 1;
    else if (result.status === "warning") summary.warnings += 1;
    else summary.clean += 1;
    return summary;
  }, { checked: 0, clean: 0, warnings: 0, errors: 0 });
}

function emitIbsherAuditProgress(run, payload = {}) {
  sendToRenderer("ibsher-audit:progress", {
    runId: run.id,
    vendorAccountValue: run.vendorAccountValue,
    processed: run.processed,
    total: run.total,
    canceled: Boolean(run.canceled),
    ...payload,
  });
}

async function requestGeminiCatalogAudit(products, settings) {
  const snapshots = products.map(ibsherProductAudit.toAiSnapshot);
  const prompt = [
    "You are auditing existing ecommerce catalog records from IBSHER.",
    "Review every product supplied. Return one audit object for every productId, even when it is clean.",
    "Check whether English, Arabic, and Central Kurdish (Sorani) names and descriptions are meaningful, mutually consistent translations, and appropriate for the same product.",
    "Arabic and Central Kurdish (Sorani) variant color names are optional in IBSHER API records. Never report a missing, empty, or untranslated Arabic/Kurdish variant color; only review those localized color values when they are present.",
    "Do not review or report collection, brand, keywords, warranty, regular or discount prices, or variant size and stock records.",
    "Check the remaining supplied details for obvious placeholder, corrupted, duplicated, contradictory, or category-mismatch values.",
    "Use deterministicIssues as confirmed structural findings. Do not invent facts, specifications, brands, categories, or errors that are not supported by the record.",
    "Use error for a definite data problem and warning for a plausible issue requiring a human review. A clean product must have an empty issues array.",
    `Products: ${JSON.stringify(snapshots)}`,
  ].join("\n");
  const schema = {
    type: "OBJECT",
    properties: {
      audits: {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          properties: {
            productId: { type: "STRING" },
            itemCode: { type: "STRING" },
            summary: { type: "STRING" },
            confidence: { type: "NUMBER" },
            issues: {
              type: "ARRAY",
              items: {
                type: "OBJECT",
                properties: {
                  field: { type: "STRING" },
                  severity: { type: "STRING", enum: ["warning", "error"] },
                  message: { type: "STRING" },
                  suggestion: { type: "STRING" },
                },
                required: ["field", "severity", "message", "suggestion"],
              },
            },
          },
          required: ["productId", "itemCode", "summary", "confidence", "issues"],
        },
      },
    },
    required: ["audits"],
  };
  const body = JSON.stringify({
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: schema,
      temperature: 0.1,
    },
  });
  const models = [DEFAULT_GEMINI_MODEL];
  let lastError;
  for (const model of models) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
    try {
      const response = await fetchWithRetry(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-goog-api-key": settings.geminiApiKey,
        },
        body,
      }, `Gemini catalog audit with ${model}`, 3);
      const data = await response.json().catch(() => ({}));
      if (response.ok) {
        return { audits: extractGeminiJson(data)?.audits || [], model };
      }
      const message = data?.error?.message || `Gemini model ${model} returned HTTP ${response.status}.`;
      if ([401, 403].includes(response.status)) throw new Error(message);
      lastError = new Error(message);
    } catch (error) {
      lastError = error;
      console.warn(`[ibsher-audit] Gemini model ${model} failed:`, error.message);
    }
  }
  throw new Error(`AI catalog audit is unavailable. ${lastError?.message || "Try again later."}`);
}

function needsGeminiCatalogAutoFix(issues = []) {
  return issues.some((issue) => {
    const field = cleanCell(issue?.field).toLowerCase();
    return field === "details"
      || field === "name"
      || field.startsWith("name.")
      || field === "description"
      || field.startsWith("description.")
      || field.includes(".color");
  });
}

async function requestGeminiCatalogAutoFix(productInput, issues, settings) {
  const product = ibsherProductAudit.normalizeIbsherProduct(productInput);
  const prompt = [
    "You are correcting one existing IBSHER ecommerce catalog record.",
    "Return corrected values only for the editable text fields in the response schema.",
    "Use the audit issues to decide what needs correction. Preserve correct existing values.",
    "Translations must describe the same product in English, Arabic, and Central Kurdish (Sorani).",
    "Collection, brand, keywords, warranty, prices, and variant size or stock records are outside this audit and must remain unchanged.",
    "Do not invent an item code, category, specification, warranty claim, or other unsupported fact.",
    "When a value cannot be corrected from the supplied record, return its existing value unchanged.",
    `Product: ${JSON.stringify({
      itemCode: product.itemCode,
      name: product.name,
      description: product.description,
      category: product.category.name,
      subCategory: product.subCategory.name,
      variants: product.variants.map((variant, variantIndex) => ({
        variantId: variant._id,
        variantIndex,
        color: variant.color,
      })),
    })}`,
    `Audit issues: ${JSON.stringify(issues)}`,
  ].join("\n");
  const localizedSchema = {
    type: "OBJECT",
    properties: {
      en: { type: "STRING" },
      ar: { type: "STRING" },
      ku: { type: "STRING" },
    },
    required: ["en", "ar", "ku"],
  };
  const schema = {
    type: "OBJECT",
    properties: {
      name: localizedSchema,
      description: localizedSchema,
      variantColors: {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          properties: {
            variantId: { type: "STRING" },
            variantIndex: { type: "NUMBER" },
            color: localizedSchema,
          },
          required: ["variantId", "variantIndex", "color"],
        },
      },
      summary: { type: "STRING" },
    },
    required: ["name", "description", "variantColors", "summary"],
  };
  const body = JSON.stringify({
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: schema,
      temperature: 0.1,
    },
  });
  const models = [DEFAULT_GEMINI_MODEL];
  let lastError;
  for (const model of models) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
    try {
      const response = await fetchWithRetry(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-goog-api-key": settings.geminiApiKey,
        },
        body,
      }, `Gemini catalog auto-fix with ${model}`, 3);
      const data = await response.json().catch(() => ({}));
      if (response.ok) return { candidate: extractGeminiJson(data), model };
      const message = data?.error?.message || `Gemini model ${model} returned HTTP ${response.status}.`;
      if ([401, 403].includes(response.status)) throw new Error(message);
      lastError = new Error(message);
    } catch (error) {
      lastError = error;
      console.warn(`[ibsher-audit] Gemini auto-fix model ${model} failed:`, error.message);
    }
  }
  throw new Error(`AI catalog auto-fix is unavailable. ${lastError?.message || "Try again later."}`);
}

function findAiAudit(aiAudits, product) {
  return aiAudits.find((audit) => cleanCell(audit.productId) === product._id)
    || aiAudits.find((audit) => cleanCell(audit.itemCode).toLowerCase() === product.itemCode.toLowerCase())
    || {
      summary: "AI did not return a review for this product.",
      confidence: 0,
      issues: [{
        field: "details",
        severity: "warning",
        message: "AI review returned no result for this product.",
        suggestion: "Run the check again.",
      }],
    };
}

async function startIbsherProductAudit(options = {}) {
  requireAccountSectionAccess("ibsher-audit");
  if (activeIbsherAudit) throw new Error("An IBSHER product audit is already running.");
  const credentials = await resolveIbsherVendorAccount(options);
  const settings = readAiSettings();
  if (!settings.geminiApiKey) {
    throw new Error("Configure and save the Gemini API key in AI Product Generator before starting an AI audit.");
  }

  const run = {
    id: `audit-${Date.now()}`,
    vendorAccountValue: credentials.value,
    canceled: false,
    processed: 0,
    total: 0,
  };
  activeIbsherAudit = run;
  const allResults = [];
  let loaded = 0;
  let model = DEFAULT_GEMINI_MODEL;
  try {
    emitIbsherAuditProgress(run, { phase: "loading", message: "Loading products from IBSHER..." });
    const scope = options.scope === "all" ? "all" : options.scope === "selected" ? "selected" : "page";
    let products;
    if (scope === "selected") {
      const seenProductIds = new Set();
      products = (Array.isArray(options.products) ? options.products : [])
        .slice(0, 500)
        .map(ibsherProductAudit.normalizeIbsherProduct)
        .filter((product) => {
          if (!product._id || seenProductIds.has(product._id)) return false;
          seenProductIds.add(product._id);
          return true;
        });
    } else if (scope === "all") {
      products = await getAllIbsherAdminProducts({ ...options, credentials }, (count, total) => {
        loaded += count;
        run.total = total;
        emitIbsherAuditProgress(run, {
          phase: "loading",
          loaded: Math.min(loaded, total),
          message: `Loaded ${Math.min(loaded, total)} of ${total} products...`,
        });
      });
    } else {
      products = (await fetchIbsherProductPage({ ...options, credentials })).products;
    }
    run.total = products.length;
    if (run.canceled) return { canceled: true, summary: getAuditSummary(allResults) };
    if (!products.length) throw new Error("No IBSHER products matched this audit.");

    const batchSize = 20;
    for (let index = 0; index < products.length; index += batchSize) {
      if (run.canceled) break;
      const batch = products.slice(index, index + batchSize);
      const ai = await requestGeminiCatalogAudit(batch, settings);
      model = ai.model;
      const results = batch.map((product) => (
        ibsherProductAudit.mergeAudit(product, findAiAudit(ai.audits, product))
      ));
      allResults.push(...results);
      run.processed = allResults.length;
      emitIbsherAuditProgress(run, {
        phase: "checking",
        model,
        results,
        summary: getAuditSummary(allResults),
        message: `AI checked ${run.processed} of ${run.total} products.`,
      });
    }

    const summary = getAuditSummary(allResults);
    emitIbsherAuditProgress(run, {
      phase: run.canceled ? "canceled" : "complete",
      model,
      summary,
      message: run.canceled
        ? `Audit stopped after ${run.processed} products.`
        : `AI audit completed for ${run.processed} products.`,
    });
    return { canceled: run.canceled, model, summary };
  } catch (error) {
    emitIbsherAuditProgress(run, { phase: "error", message: error.message });
    throw error;
  } finally {
    if (activeIbsherAudit?.id === run.id) activeIbsherAudit = null;
  }
}

function cancelIbsherProductAudit() {
  requireAccountSectionAccess("ibsher-audit");
  if (!activeIbsherAudit) return { canceled: false, running: false };
  activeIbsherAudit.canceled = true;
  return { canceled: true, running: true, runId: activeIbsherAudit.id };
}

async function autoFixIbsherAdminProduct(options = {}) {
  requireAccountSectionAccess("ibsher-audit");
  if (activeIbsherAudit) {
    throw new Error("Stop the running catalog audit before automatically fixing a product.");
  }
  const credentials = await resolveIbsherVendorAccount(options);

  const suppliedProduct = ibsherProductAudit.normalizeIbsherProduct(options.product);
  if (!suppliedProduct._id) throw new Error("The IBSHER product ID is missing.");

  let product = suppliedProduct;
  if (suppliedProduct.itemCode) {
    const latestPage = await fetchIbsherProductPage({
      search: suppliedProduct.itemCode,
      limit: 50,
      credentials,
    });
    product = latestPage.products.find((entry) => entry._id === suppliedProduct._id) || suppliedProduct;
  }

  const suppliedIssues = Array.isArray(options.issues) ? options.issues.slice(0, 100) : [];
  const issues = suppliedIssues.length
    ? suppliedIssues.map((issue) => ({
      field: cleanCell(issue?.field).slice(0, 160),
      severity: cleanCell(issue?.severity).toLowerCase() === "error" ? "error" : "warning",
      message: cleanCell(issue?.message).slice(0, 1000),
      suggestion: cleanCell(issue?.suggestion).slice(0, 1000),
    }))
    : ibsherProductAudit.deterministicAudit(product).issues;
  const settings = readAiSettings();
  let candidate = {};
  let model = "";

  if (needsGeminiCatalogAutoFix(issues)) {
    if (!settings.geminiApiKey) {
      throw new Error("Configure and save the Gemini API key before using automatic catalog fixes.");
    }
    const aiFix = await requestGeminiCatalogAutoFix(product, issues, settings);
    candidate = aiFix.candidate;
    model = aiFix.model;
  }

  const fixed = ibsherProductAudit.applySafeAutoFix(product, candidate, issues);
  if (!fixed.changes.length) {
    return {
      saved: false,
      product,
      changes: [],
      audit: ibsherProductAudit.mergeAudit(product, {
        summary: "No safe automatic correction was available. Edit this product manually.",
        confidence: 1,
        issues,
      }),
    };
  }

  const remainingValidationErrors = getIbsherProductUpdateValidationErrors(fixed.product);
  if (remainingValidationErrors.length) {
    return {
      saved: false,
      product,
      changes: [],
      audit: ibsherProductAudit.mergeAudit(product, {
        summary: "Automatic corrections could not be saved because this product still needs manual fixes.",
        confidence: 1,
        issues: [
          ...issues,
          ...remainingValidationErrors.slice(0, 12).map((message) => ({
            field: "details",
            severity: "error",
            message,
            suggestion: "Edit this field manually, then run the audit again.",
          })),
        ],
      }),
    };
  }

  const updated = await updateIbsherAdminProductWithCredentials(fixed.product, credentials);
  const savedProduct = updated.product || fixed.product;
  let audit;
  if (settings.geminiApiKey) {
    try {
      const recheck = await requestGeminiCatalogAudit([savedProduct], settings);
      model = recheck.model || model;
      audit = ibsherProductAudit.mergeAudit(savedProduct, findAiAudit(recheck.audits, savedProduct));
    } catch (error) {
      audit = ibsherProductAudit.mergeAudit(savedProduct, {
        summary: "Automatic changes were saved, but AI verification could not be completed.",
        confidence: 0,
        issues: [{
          field: "details",
          severity: "warning",
          message: "The saved automatic changes still need an AI recheck.",
          suggestion: "Run the catalog audit again for this product.",
        }],
      });
    }
  } else {
    audit = ibsherProductAudit.mergeAudit(savedProduct, {
      summary: "Automatic deterministic corrections were saved.",
      confidence: 1,
      issues: [],
    });
  }

  return {
    saved: true,
    product: savedProduct,
    audit,
    changes: fixed.changes,
    model,
  };
}

function getIbsherProductUpdateValidationErrors(productInput = {}, options = {}) {
  const product = ibsherProductAudit.normalizeIbsherProduct(productInput);
  const replacementImagesByVariant = options.replacementImagesByVariant || {};
  const errors = [];
  const checkLength = (field, value, minimum, maximum) => {
    const length = cleanCell(value).length;
    if (length < minimum) errors.push(`${field} must contain at least ${minimum} characters.`);
    else if (length > maximum) errors.push(`${field} must not exceed ${maximum} characters.`);
  };

  for (const language of ["en", "ar", "ku"]) {
    checkLength(`name.${language}`, product.name[language], 2, 100);
    checkLength(`description.${language}`, product.description[language], 10, 3000);
  }
  if (!product.itemCode) errors.push("itemCode is required.");
  if (!product.collectionName._id) errors.push("collectionName is required.");
  if (!product.brand._id) errors.push("brand is required.");
  if (!(product.price > 0)) errors.push("price must be greater than zero.");
  if (product.discountPrice < 0 || (product.discountPrice > 0 && product.discountPrice >= product.price)) {
    errors.push("discountPrice must be zero or lower than price.");
  }
  if (!product.keyword.length || product.keyword.some((keyword) => !cleanCell(keyword))) {
    errors.push("At least one non-empty keyword is required.");
  }
  if (product.rating < 0 || product.rating > 5) errors.push("rating must be between 0 and 5.");
  if (!options.skipRewardValidation) {
    if (product.points < 0) errors.push("points cannot be negative.");
    if (product.cashback < 0) errors.push("cashback cannot be negative.");
  }

  const hasWarranty = product.warranty.duration > 0
    || Boolean(product.warranty.unit)
    || Boolean(product.warranty.description);
  if (hasWarranty && !["days", "weeks", "months", "years"].includes(product.warranty.unit)) {
    errors.push("warranty.unit must be days, weeks, months, or years.");
  }
  if (!product.variants.length) errors.push("At least one variant is required.");
  const colors = new Set();
  product.variants.forEach((variant, variantIndex) => {
    checkLength(`variants.${variantIndex}.color.en`, variant.color.en, 1, 40);
    for (const language of ["ar", "ku"]) {
      if (variant.color[language].length > 40) {
        errors.push(`variants.${variantIndex}.color.${language} must not exceed 40 characters.`);
      }
    }
    const colorKey = cleanCell(variant.color.en).toLowerCase();
    if (colorKey && colors.has(colorKey)) errors.push(`variants.${variantIndex}.color.en duplicates another variant color.`);
    colors.add(colorKey);
    const replacementImages = replacementImagesByVariant instanceof Map
      ? replacementImagesByVariant.get(variantIndex)
      : replacementImagesByVariant[variantIndex];
    if (!variant.images.length && !replacementImages?.length) {
      errors.push(`variants.${variantIndex} requires at least one existing or replacement image.`);
    }
    if (!variant.sizes.length) errors.push(`variants.${variantIndex} requires at least one size.`);
    const sizes = new Set();
    variant.sizes.forEach((size, sizeIndex) => {
      const prefix = `variants.${variantIndex}.sizes.${sizeIndex}`;
      checkLength(`${prefix}.size`, size.size, 1, 50);
      const sizeKey = cleanCell(size.size).toLowerCase();
      if (sizeKey && sizes.has(sizeKey)) errors.push(`${prefix}.size duplicates another size in this variant.`);
      sizes.add(sizeKey);
      if (!["in_stock", "out_of_stock"].includes(size.stockStatus)) errors.push(`${prefix}.stockStatus is invalid.`);
      if (!Number.isInteger(size.stockQuantity) || size.stockQuantity < 0) errors.push(`${prefix}.stockQuantity must be a non-negative integer.`);
      if (size.stockStatus === "in_stock" && size.stockQuantity <= 0) errors.push(`${prefix}.stockQuantity must be greater than zero for in_stock.`);
      if (size.stockStatus === "out_of_stock" && size.stockQuantity !== 0) errors.push(`${prefix}.stockQuantity must be zero for out_of_stock.`);
    });
  });
  return [...new Set(errors)];
}

function createIbsherProductUpdateForm(productInput = {}, options = {}) {
  const product = ibsherProductAudit.normalizeIbsherProduct(productInput);
  if (!product._id) throw new Error("The IBSHER product ID is missing.");
  const replacementImagesByVariant = options.replacementImagesByVariant || {};
  const validationErrors = getIbsherProductUpdateValidationErrors(product, {
    replacementImagesByVariant, skipRewardValidation: options.skipRewardValidation === true,
  });
  if (validationErrors.length) {
    throw new Error(`IBSHER product validation failed: ${validationErrors.slice(0, 8).join("; ")}`);
  }

  const form = new FormData();
  form.append("name", JSON.stringify(product.name));
  form.append("description", JSON.stringify(product.description));
  form.append("itemCode", product.itemCode);
  form.append("keyword", JSON.stringify(product.keyword));
  for (const field of ["category", "subCategory", "collectionName", "brand"]) {
    if (product[field]?._id) form.append(field, product[field]._id);
  }
  form.append("currency", product.currency === "USD" ? "USD" : "IQD");
  form.append("price", String(product.price));
  form.append("discountPrice", String(product.discountPrice || 0));
  const hasWarranty = product.warranty.duration > 0
    || Boolean(product.warranty.unit)
    || Boolean(product.warranty.description);
  if (hasWarranty) {
    form.append("warranty", JSON.stringify({
      duration: product.warranty.duration,
      unit: product.warranty.unit,
      description: product.warranty.description || undefined,
    }));
  }
  form.append("isActive", String(product.isActive));
  form.append("is_featured", String(product.is_featured));
  form.append("is_new_arrival", String(product.is_new_arrival));
  form.append("is_hot", String(product.is_hot));
  form.append("is_best_seller", String(product.is_best_seller));
  form.append("badge", product.badge);
  form.append("rating", String(product.rating));
  form.append("variants", JSON.stringify(product.variants.map((variant, variantIndex) => {
    const replacementImages = replacementImagesByVariant instanceof Map
      ? replacementImagesByVariant.get(variantIndex)
      : replacementImagesByVariant[variantIndex];
    const replacingImages = Array.isArray(replacementImages) && replacementImages.length > 0;
    return {
      ...(variant._id ? { _id: variant._id } : {}),
      color: variant.color,
      existingImages: replacingImages ? [] : variant.images.map((image) => ({
        _id: image._id,
        url: image.url,
        key: image.key,
        alt: image.alt,
        isMain: image.isMain,
      })),
      removedImages: replacingImages ? variant.images.map((image) => ({
        _id: image._id,
        url: image.url,
      })) : [],
      sizes: variant.sizes.map((size) => ({
        ...(size._id ? { _id: size._id } : {}),
        size: cleanCell(size.size),
        stockStatus: size.stockStatus || "in_stock",
        stockQuantity: size.stockStatus === "out_of_stock" ? 0 : Number(size.stockQuantity || 0),
      })),
    };
  })));
  product.variants.forEach((variant, variantIndex) => {
    const replacementImages = replacementImagesByVariant instanceof Map
      ? replacementImagesByVariant.get(variantIndex)
      : replacementImagesByVariant[variantIndex];
    if (!Array.isArray(replacementImages)) return;
    replacementImages.forEach((imagePath, imageIndex) => {
      form.append(
        `variants[${variantIndex}][images][${imageIndex}]`,
        fs.createReadStream(imagePath),
      );
    });
  });
  return { form, product };
}

async function updateIbsherAdminProductWithCredentials(productInput = {}, credentialsInput = {}, options = {}) {
  const credentials = getIbsherAdminCredentials(credentialsInput);
  const normalized = ibsherProductAudit.normalizeIbsherProduct(productInput);
  const response = await requestIbsherAdmin(() => {
    const { form, product } = createIbsherProductUpdateForm(normalized, options);
    return {
      method: "PUT",
      url: `/product/${encodeURIComponent(product._id)}`,
      data: form,
      headers: form.getHeaders(),
      maxContentLength: Infinity,
      maxBodyLength: Infinity,
    };
  }, credentials);
  const refreshed = await fetchIbsherProductPage({
    search: normalized.itemCode,
    limit: 50,
    credentials,
  });
  const returnedProduct = response.data?.product || response.data?.data;
  return {
    product: refreshed.products.find((product) => product._id === normalized._id)
      || refreshed.products[0]
      || (returnedProduct && typeof returnedProduct === "object"
        ? addIbsherProductDisplayUrls(returnedProduct)
        : null)
      || normalized,
  };
}

async function updateIbsherAdminProduct(productInput = {}, options = {}) {
  requireAccountSectionAccess("ibsher-audit");
  const credentials = await resolveIbsherVendorAccount(options);
  return updateIbsherAdminProductWithCredentials(productInput, credentials);
}

function splitImageCell(value) {
  return cleanCell(value)
    .split(/[\n\r,;|]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function getRowImageValues(normalizedEntries) {
  const bulkImages = [
    ...splitImageCell(normalizedEntries.images),
    ...splitImageCell(normalizedEntries.image),
  ];
  const numberedImages = Object.entries(normalizedEntries)
    .filter(([key]) => /^image-\d+$/.test(key))
    .sort(([left], [right]) => Number(left.slice(6)) - Number(right.slice(6)))
    .flatMap(([, value]) => splitImageCell(value));
  return [...bulkImages, ...numberedImages];
}

function getImageMimeType(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  return {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".heic": "image/heic",
    ".heif": "image/heif",
  }[extension] || "";
}

const AI_PRODUCT_IMAGE_REENCODE_BYTES = 1.5 * 1024 * 1024;
const AI_PRODUCT_JPEG_QUALITY = 86;

function optimizeAiProductImage(buffer, mimeType) {
  if (!Buffer.isBuffer(buffer) || !buffer.length) return { buffer, mimeType };

  try {
    const image = nativeImage.createFromBuffer(buffer);
    if (image.isEmpty()) return { buffer, mimeType };
    const dimensions = aiProductPerformance.getConstrainedImageDimensions(
      image.getSize().width,
      image.getSize().height
    );
    const isPng = mimeType === "image/png";
    const shouldReencode = dimensions.resized || (!isPng && buffer.length > AI_PRODUCT_IMAGE_REENCODE_BYTES);
    if (!shouldReencode) return { buffer, mimeType };

    const preparedImage = dimensions.resized
      ? image.resize({ width: dimensions.width, height: dimensions.height, quality: "best" })
      : image;
    const optimizedBuffer = isPng
      ? preparedImage.toPNG()
      : preparedImage.toJPEG(AI_PRODUCT_JPEG_QUALITY);
    if (!optimizedBuffer.length || (!dimensions.resized && optimizedBuffer.length >= buffer.length)) {
      return { buffer, mimeType };
    }

    return {
      buffer: optimizedBuffer,
      mimeType: isPng ? "image/png" : "image/jpeg",
    };
  } catch (error) {
    console.warn(`[ai-product] Could not optimize an image payload: ${error.message}`);
    return { buffer, mimeType };
  }
}

const AI_PRODUCT_IMAGE_EXTENSIONS = new Set([
  ".jpg",
  ".jpeg",
  ".png",
  ".webp",
  ".heic",
  ".heif",
]);

async function createAiImageInfo(filePath, options = {}) {
  const mimeType = getImageMimeType(filePath);
  if (!mimeType || !fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) return null;
  const stats = fs.statSync(filePath);
  if (options.thumbnail && typeof nativeImage.createThumbnailFromPath === "function") {
    try {
      const systemThumbnail = await nativeImage.createThumbnailFromPath(filePath, { width: 180, height: 180 });
      if (!systemThumbnail.isEmpty()) {
        const thumbnailBuffer = systemThumbnail.toJPEG(78);
        if (thumbnailBuffer.length) {
          return {
            path: filePath,
            name: path.basename(filePath),
            size: stats.size,
            mimeType,
            previewUrl: `data:image/jpeg;base64,${thumbnailBuffer.toString("base64")}`,
          };
        }
      }
    } catch (error) {
      console.warn(`[ai-product] System thumbnail failed for ${path.basename(filePath)}: ${error.message}`);
    }
  }
  const buffer = fs.readFileSync(filePath);
  let previewImage = await prepareAiProductImageBuffer(buffer, mimeType, filePath);
  if (options.thumbnail) {
    const image = nativeImage.createFromBuffer(previewImage.buffer);
    if (!image.isEmpty()) {
      const size = image.getSize();
      const scale = Math.min(1, 180 / Math.max(size.width, size.height));
      const thumbnail = scale < 1
        ? image.resize({ width: Math.max(1, Math.round(size.width * scale)), quality: "best" })
        : image;
      const thumbnailBuffer = thumbnail.toJPEG(82);
      if (thumbnailBuffer.length) previewImage = { buffer: thumbnailBuffer, mimeType: "image/jpeg" };
    }
  }
  return {
    path: filePath,
    name: path.basename(filePath),
    size: stats.size,
    mimeType,
    previewUrl: `data:${previewImage.mimeType};base64,${previewImage.buffer.toString("base64")}`,
  };
}

function isRemoteImageSource(source) {
  return /^https?:\/\//i.test(String(source || "").trim());
}

function createRemoteAiImageInfo(imageUrl) {
  try {
    const parsed = new URL(imageUrl);
    return {
      path: imageUrl,
      name: path.basename(parsed.pathname) || "remote-image",
      size: 0,
      mimeType: getImageMimeType(parsed.pathname),
      previewUrl: imageUrl,
      isRemote: true,
    };
  } catch {
    return null;
  }
}

async function fetchWithRetry(requestUrl, options = {}, label = "Network request", maxAttempts = 3) {
  let lastError;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = await net.fetch(requestUrl, options);
      if (response.status !== 429 && response.status < 500) return response;
      lastError = new Error(`${label} returned HTTP ${response.status}.`);
    } catch (error) {
      lastError = error;
    }
    if (attempt < maxAttempts) {
      await new Promise((resolve) => setTimeout(resolve, Math.min(8000, 1000 * (2 ** (attempt - 1)))));
    }
  }
  throw new Error(`${label} failed after ${maxAttempts} attempts: ${lastError?.message || "Network unavailable."}`);
}

const PRODUCT_SEARCH_HEADERS = {
  Accept: "text/html,application/xhtml+xml,application/json;q=0.8,*/*;q=0.5",
  "Accept-Language": "en-US,en;q=0.9,ar;q=0.7",
  "Cache-Control": "no-cache",
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
};

async function fetchProductSearchPage(pageUrl, label = "Product website") {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 12000);
  try {
    const response = await fetchWithRetry(pageUrl, {
      headers: PRODUCT_SEARCH_HEADERS,
      redirect: "follow",
      signal: controller.signal,
    }, label, 1);
    if (!response.ok) return null;
    const contentType = String(response.headers.get("content-type") || "").toLowerCase();
    if (contentType && !/(?:text\/html|application\/xhtml\+xml|application\/json|text\/plain)/.test(contentType)) {
      return null;
    }
    const html = (await response.text()).slice(0, 8 * 1024 * 1024);
    return { url: response.url || pageUrl, html };
  } catch (error) {
    console.warn(`[product-image-search] ${label} failed for ${pageUrl}: ${error.message}`);
    return null;
  } finally {
    clearTimeout(timeout);
  }
}

async function mapWithConcurrency(values, concurrency, mapper) {
  const results = new Array(values.length);
  let nextIndex = 0;
  const workers = Array.from({ length: Math.min(concurrency, values.length) }, async () => {
    while (nextIndex < values.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await mapper(values[currentIndex], currentIndex);
    }
  });
  await Promise.all(workers);
  return results;
}

async function discoverProductPagesWithWebSearch(product, websites) {
  const queries = productImageSearch.buildSearchQueries(product, websites);
  const searchResults = await mapWithConcurrency(queries, 3, async (query) => {
    const searchUrl = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`;
    const page = await fetchProductSearchPage(searchUrl, "Product web search");
    return page ? productImageSearch.extractPageLinks(page.html, page.url, websites) : [];
  });
  return [...new Set(searchResults.flat())];
}

async function discoverProductPagesWithSiteSearch(product, websites) {
  const searchUrls = productImageSearch.buildSiteSearchUrls(product, websites);
  const searchResults = await mapWithConcurrency(searchUrls, 3, async (searchUrl) => {
    const page = await fetchProductSearchPage(searchUrl, "Official site search");
    if (!page) return [];
    return productImageSearch.extractPageLinks(page.html, page.url, websites);
  });
  return [...new Set(searchResults.flat())];
}

function extractHttpUrls(value) {
  return [...String(value || "").matchAll(/https?:\/\/[^\s<>"'`]+/gi)]
    .map((match) => match[0].replace(/[),.;\]}]+$/g, ""));
}

async function resolveGroundingUrl(candidateUrl) {
  try {
    const response = await net.fetch(candidateUrl, {
      headers: PRODUCT_SEARCH_HEADERS,
      redirect: "follow",
    });
    const resolvedUrl = response.url || candidateUrl;
    response.body?.cancel?.().catch(() => {});
    return resolvedUrl;
  } catch {
    return candidateUrl;
  }
}

async function discoverProductPagesWithGemini(product, websites) {
  const settings = readAiSettings();
  if (!settings.geminiApiKey) return [];
  const prompt = [
    "Use Google Search to find the single best official manufacturer product-detail page for this exact retail item.",
    `Item code: ${product.itemcode}`,
    `Brand: ${product.brand || "not supplied"}`,
    `English item name: ${product.en || "not supplied"}`,
    `Allowed official websites: ${websites.map((website) => website.url).join(", ")}`,
    "The ordered product gallery images from this page will be uploaded to an e-commerce catalog.",
    "Choose pages whose gallery shows only the exact product/packshot views, not description, feature, lifestyle, accessory, review, or related-product content.",
    "Return only complete product-detail page URLs, one URL per line. Do not return retailer, marketplace, PDF, support, category, or search-result URLs.",
  ].join("\n");
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(FREE_SEARCH_GEMINI_MODEL)}:generateContent`;
  try {
    const response = await fetchWithRetry(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": settings.geminiApiKey,
      },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        tools: [{ google_search: {} }],
        generationConfig: { temperature: 0 },
      }),
    }, "AI product page search", 2);
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      console.warn(`[product-image-search] Gemini search returned HTTP ${response.status}.`);
      return [];
    }
    const responseText = data?.candidates?.[0]?.content?.parts?.map((part) => part.text || "").join("\n") || "";
    const groundingUrls = (data?.candidates?.[0]?.groundingMetadata?.groundingChunks || [])
      .map((chunk) => chunk?.web?.uri)
      .filter(Boolean);
    const rawUrls = [...new Set([...extractHttpUrls(responseText), ...groundingUrls])];
    const resolvedUrls = await mapWithConcurrency(rawUrls, 3, async (candidateUrl) => (
      productImageSearch.isAllowedPageUrl(candidateUrl, websites)
        ? candidateUrl
        : resolveGroundingUrl(candidateUrl)
    ));
    return [...new Set(resolvedUrls.filter((candidateUrl) => productImageSearch.isAllowedPageUrl(candidateUrl, websites)))];
  } catch (error) {
    console.warn(`[product-image-search] Gemini page discovery failed: ${error.message}`);
    return [];
  }
}

function selectProductPageCandidates(candidateUrls, product, websites) {
  const perHost = new Map();
  return [...new Set(candidateUrls)]
    .filter((candidateUrl) => productImageSearch.isAllowedPageUrl(candidateUrl, websites))
    .sort((left, right) => (
      productImageSearch.scoreProductPageUrl(right, product)
      - productImageSearch.scoreProductPageUrl(left, product)
    ))
    .filter((candidateUrl) => {
      const hostname = new URL(candidateUrl).hostname;
      const count = perHost.get(hostname) || 0;
      if (count >= 3) return false;
      perHost.set(hostname, count + 1);
      return true;
    })
    .slice(0, 18);
}

async function extractImagesFromProductPages(candidateUrls, product, websites, options = {}) {
  const selectedUrls = selectProductPageCandidates(candidateUrls, product, websites);
  const pages = await mapWithConcurrency(selectedUrls, 4, async (pageUrl) => (
    fetchProductSearchPage(pageUrl, `Official page for ${product.itemcode}`)
  ));
  const extractedPages = await mapWithConcurrency(pages, 3, async (page, candidateOrder) => {
    if (!page) return null;
    if (!productImageSearch.isAllowedPageUrl(page.url, websites)) return null;
    const matchesProduct = options.requireExactItemCode
      ? productImageSearch.pageMatchesItemCode(page.html, page.url, product.itemcode)
      : productImageSearch.pageMatchesProduct(page.html, page.url, product);
    if (!matchesProduct) return null;
    const adapter = productImageSearch.detectProductPageAdapter(page.url, page.html);
    let pageImages = [];
    const productDataUrls = productImageSearch.buildProductImageDataUrls(page.url, adapter);
    for (const productDataUrl of productDataUrls) {
      const productDataPage = await fetchProductSearchPage(productDataUrl, `${adapter} product gallery data`);
      if (!productDataPage) continue;
      pageImages = productImageSearch.extractProductImagesFromData(
        productDataPage.html,
        page.url,
        adapter
      );
      if (pageImages.length) break;
    }
    if (!pageImages.length) {
      pageImages = productImageSearch.extractProductImages(page.html, page.url, product);
    }
    if (!pageImages.length) return null;
    return { url: page.url, images: pageImages, candidateOrder, adapter };
  });
  const pageResults = extractedPages.filter(Boolean);
  const selectedPage = productImageSearch.chooseBestProductPageResult(pageResults, product, websites);
  return selectedPage
    ? { images: selectedPage.images, productPages: [selectedPage.url], adapter: selectedPage.adapter }
    : { images: [], productPages: [], adapter: "" };
}

async function findImagesForProduct(product, websites) {
  const directUrls = productImageSearch.buildDirectProductUrls(product, websites);
  const [webSearchUrls, aiSearchUrls] = await Promise.all([
    discoverProductPagesWithWebSearch(product, websites),
    discoverProductPagesWithGemini(product, websites),
  ]);
  let extracted = await extractImagesFromProductPages(
    [...directUrls, ...aiSearchUrls, ...webSearchUrls],
    product,
    websites
  );
  if (!extracted.images.length) {
    const siteSearchUrls = await discoverProductPagesWithSiteSearch(product, websites);
    extracted = await extractImagesFromProductPages(
      [...directUrls, ...siteSearchUrls],
      product,
      websites
    );
  }
  return extracted;
}

function isUsableDiscoveredProductPage(candidateUrl) {
  try {
    const parsed = new URL(candidateUrl);
    if (!/^https?:$/.test(parsed.protocol)) return false;
    if (/\.(?:jpe?g|png|webp|avif|gif|svg|pdf)(?:$|[?#])/i.test(parsed.pathname)) return false;
    return !/(?:google\.|gstatic\.|duckduckgo\.|bing\.|youtube\.|youtu\.be$)/i.test(parsed.hostname);
  } catch {
    return false;
  }
}

async function discoverExactProductPagesWithGemini(product, settings) {
  const prompt = [
    "Use Google Search to identify the exact retail product matching this item code.",
    `Item code: ${product.itemcode}`,
    `Brand: ${product.brand || "not supplied"}`,
    `English product name: ${product.en || "not supplied"}`,
    "Return up to five complete product-detail page URLs, with the official manufacturer page first when available.",
    "Every returned page must visibly contain the exact item code and show the product gallery for that exact item.",
    "Do not return search results, category pages, PDFs, videos, social media, reviews, related products, accessories, or image URLs.",
    "Output only one complete URL per line.",
  ].join("\n");
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(FREE_SEARCH_GEMINI_MODEL)}:generateContent`;
  const response = await fetchWithRetry(endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-goog-api-key": settings.geminiApiKey,
    },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      tools: [{ google_search: {} }],
      generationConfig: { temperature: 0 },
    }),
  }, `AI image search for ${product.itemcode}`, 2);
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data?.error?.message || `Image search returned HTTP ${response.status}.`);
  }
  const responseText = data?.candidates?.[0]?.content?.parts
    ?.map((part) => part.text || "")
    .join("\n") || "";
  const groundingUrls = (data?.candidates?.[0]?.groundingMetadata?.groundingChunks || [])
    .map((chunk) => chunk?.web?.uri)
    .filter(Boolean);
  const rawUrls = [...new Set([...extractHttpUrls(responseText), ...groundingUrls])];
  const resolvedUrls = await mapWithConcurrency(rawUrls, 3, resolveGroundingUrl);
  return [...new Set(resolvedUrls.filter(isUsableDiscoveredProductPage))].slice(0, 12);
}

async function findExactProductImageForAiRow(product, settings) {
  const itemcode = cleanCell(product.itemcode);
  if (!itemcode) {
    throw new Error("An item code is required to find a missing product image.");
  }
  const normalizedProduct = {
    itemcode,
    brand: cleanCell(product.brand),
    en: cleanCell(product.en),
  };
  const candidateUrls = await discoverExactProductPagesWithGemini(normalizedProduct, settings);
  if (!candidateUrls.length) {
    throw new Error(`No exact product page was found for item code ${itemcode}.`);
  }
  const websites = productImageSearch.normalizeWebsites(
    [...new Set(candidateUrls.map((candidateUrl) => new URL(candidateUrl).origin))]
  );
  const found = await extractImagesFromProductPages(
    candidateUrls,
    normalizedProduct,
    websites,
    { requireExactItemCode: true }
  );
  const imageUrl = cleanCell(found.images?.[0]);
  if (!imageUrl) {
    throw new Error(`No verified product-only image was found for item code ${itemcode}.`);
  }
  return imageUrl;
}

async function findOfficialProductImages(options = {}) {
  requireAuthenticatedUser();
  const products = productImageSearch.normalizeLookupProducts(options.products);
  const websites = productImageSearch.normalizeWebsites(options.websites);
  const results = [];
  for (let index = 0; index < products.length; index += 1) {
    const product = products[index];
    sendToRenderer("product-image-search:progress", {
      current: index + 1,
      total: products.length,
      itemcode: product.itemcode,
    });
    try {
      const found = await findImagesForProduct(product, websites);
      results.push({
        ...product,
        ...found,
        error: found.images.length ? "" : "No matching product images were found on the selected websites.",
      });
    } catch (error) {
      console.warn(`[product-image-search] ${product.itemcode} failed: ${error.message}`);
      results.push({ ...product, images: [], productPages: [], error: error.message });
    }
  }
  return { results };
}

async function importProductImageSearchExcel() {
  requireAuthenticatedUser();
  const selection = await dialog.showOpenDialog(mainWindow, {
    title: "Choose product codes and brands Excel",
    properties: ["openFile"],
    filters: [
      { name: "Excel Workbook", extensions: ["xlsx", "xls", "csv"] },
    ],
  });
  if (selection.canceled || !selection.filePaths[0]) {
    return { canceled: true, products: [], skipped: [] };
  }
  const filePath = selection.filePaths[0];
  const workbook = loadXlsx().readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("The selected Excel workbook has no sheets.");
  const rows = loadXlsx().utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  const parsed = productImageSearch.parseProductImageWorkbookRows(rows);
  if (!parsed.products.length) {
    throw new Error("No valid products were found. Excel requires itemcode and brand columns; en is optional.");
  }
  return {
    canceled: false,
    filePath,
    sheetName,
    products: parsed.products,
    skipped: parsed.skipped,
  };
}

async function exportProductImageExcel(options = {}) {
  requireAuthenticatedUser();
  const { columns, sourceRows, excelRows } = productImageSearch.buildProductImageExcelData(options.rows);
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Download product image links Excel",
    defaultPath: `product-images-${Date.now()}.xlsx`,
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (result.canceled || !result.filePath) return { canceled: true };
  const xlsx = loadXlsx();
  const worksheet = xlsx.utils.json_to_sheet(excelRows, { header: columns });
  worksheet["!cols"] = columns.map((column) => ({ wch: column === "itemcode" ? 24 : 60 }));
  for (let rowIndex = 0; rowIndex < excelRows.length; rowIndex += 1) {
    for (let imageIndex = 0; imageIndex < sourceRows[rowIndex].images.length; imageIndex += 1) {
      const imageColumnIndex = columns.indexOf(`image-${imageIndex + 1}`);
      const cellAddress = xlsx.utils.encode_cell({ r: rowIndex + 1, c: imageColumnIndex });
      if (worksheet[cellAddress]) worksheet[cellAddress].l = { Target: sourceRows[rowIndex].images[imageIndex] };
    }
  }
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Product Images");
  xlsx.writeFile(workbook, result.filePath);
  return { canceled: false, filePath: result.filePath };
}

let activeBrandImport = null;

function parseIbsherBrandWorkbook(options = {}) {
  requireAuthenticatedUser();
  const filePath = path.resolve(cleanCell(options.filePath));
  if (!cleanCell(options.filePath) || !/\.(xlsx|xls|csv)$/i.test(filePath)) {
    throw new Error("Choose an Excel or CSV file for brand import.");
  }
  const xlsx = loadXlsx();
  const workbook = xlsx.readFile(filePath);
  const sheetName = cleanCell(options.sheetName) || workbook.SheetNames[0];
  if (!sheetName || !workbook.Sheets[sheetName]) throw new Error("The selected Excel sheet was not found.");
  let parsed;
  let parseError = "";
  try {
    parsed = ibsherBrandImport.parseBrandWorksheet(xlsx, workbook.Sheets[sheetName]);
  } catch (error) {
    parsed = { rows: [], validCount: 0, invalidCount: 0 };
    parseError = error.message;
  }
  return { canceled: false, filePath, sheetName, sheetNames: workbook.SheetNames, parseError, ...parsed };
}

let activeProductPublishing = null;
let activeIntegrationReview = null;
let activeVendorProductDeletion = null;

async function selectIbsherProductDeletionExcel() {
  requireAccountSectionAccess("integration-review");
  const selection = await dialog.showOpenDialog(mainWindow, {
    title: "Choose the Excel file containing product SKUs",
    properties: ["openFile"],
    filters: [{ name: "Excel Workbooks", extensions: ["xlsx", "xls", "csv"] }],
  });
  if (selection.canceled || !selection.filePaths?.[0]) return { canceled: true };
  const filePath = selection.filePaths[0];
  const xlsx = loadXlsx();
  const workbook = xlsx.readFile(filePath, { raw: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("The selected Excel workbook has no worksheet.");
  const matrix = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
  const parsed = ibsherVendorProductDelete.parseSkuMatrix(matrix);
  return {
    canceled: false,
    filePath,
    fileName: path.basename(filePath),
    sheetName,
    skus: parsed.skus,
    rows: parsed.rows,
    skipped: parsed.skipped,
  };
}

async function collectIbsherDeletionCatalog(credentials, progress = () => {}, shouldCancel = () => false) {
  return ibsherVendorProductDelete.collectCompleteProductCatalog({
    listProducts: (pageOptions) => fetchIbsherProductApiPage({ ...pageOptions, credentials, requireComplete: true }),
    shouldCancel,
    onProgress: progress,
  });
}

async function previewIbsherVendorProductDeletion(options = {}) {
  requireAccountSectionAccess("integration-review");
  const credentials = await resolveIbsherVendorAccount(options);
  const vendorAccountValue = getRequestedVendorValue(options);
  const products = await collectIbsherDeletionCatalog(credentials, (progress) => (
    sendToRenderer("ibsher-product-delete:progress", { ...progress, vendorAccountValue })
  ));
  return ibsherVendorProductDelete.buildDeletionPreview(options.skus, products);
}

async function deleteIbsherVendorProductsFromExcel(options = {}) {
  requireAccountSectionAccess("integration-review");
  if (activeVendorProductDeletion) throw new Error("An Excel product deletion is already running.");
  const run = { canceled: false, userValue: authSession?.user?.value };
  activeVendorProductDeletion = run;
  try {
    const credentials = await resolveIbsherVendorAccount(options);
    const vendorAccountValue = getRequestedVendorValue(options);
    return await ibsherVendorProductDelete.deleteProductsBySku(options, {
      listProducts: (pageOptions) => fetchIbsherProductApiPage({ ...pageOptions, credentials, requireComplete: true }),
      deleteProduct: async (productId) => (await requestIbsherAdmin({
        method: "DELETE", url: `/product/${encodeURIComponent(productId)}`,
      }, credentials)).data,
      shouldCancel: () => run.canceled || authSession?.user?.value !== run.userValue,
      onProgress: (progress) => sendToRenderer("ibsher-product-delete:progress", { ...progress, vendorAccountValue }),
    });
  } finally {
    activeVendorProductDeletion = null;
  }
}

async function fetchIbsherVendorImportPage(options = {}, credentials) {
  const page = Math.max(1, Number.parseInt(options.page, 10) || 1);
  const limit = Math.max(1, Math.min(100, Number.parseInt(options.limit, 10) || 100));
  const status = cleanCell(options.status);
  const search = cleanCell(options.search).slice(0, 160);
  const response = await requestIbsherAdmin({
    method: "GET",
    url: "/vendor-product-imports",
    params: { page, limit, ...(status ? { status } : {}), ...(search ? { search } : {}) },
  }, credentials);
  if (response.data?.success === false || !Array.isArray(response.data?.products)) {
    throw new Error(cleanCell(response.data?.message) || "IBSHER returned an invalid vendor import list.");
  }
  return ibsherIntegrationReview.normalizeImportPage(response.data, { page, limit });
}

async function loadIbsherVendorImport(importId, credentials) {
  const id = ibsherIntegrationReview.validateImportId(importId);
  const response = await requestIbsherAdmin({
    method: "GET", url: `/vendor-product-imports/${encodeURIComponent(id)}`,
  }, credentials);
  const product = response.data?.product;
  if (response.data?.success === false || !product) {
    throw new Error(cleanCell(response.data?.message) || "IBSHER returned an invalid vendor product import.");
  }
  return product;
}

async function loadIbsherIntegrationCatalog(credentials) {
  const [categoryResponse, subcategories, collections, brands] = await Promise.all([
    requestIbsherAdmin({ method: "GET", url: "/category" }, credentials),
    fetchIbsherAdminListPages("/subcategory", credentials, [
      "subCategory", "subCategories", "subcategory", "subcategories", "data",
    ], { label: "subcategory" }),
    fetchIbsherAdminListPages("/collection", credentials, [
      "collections", "collection", "data",
    ], { label: "collection" }),
    fetchIbsherAdminListPages("/brand", credentials, ["brands", "brand", "data"], { label: "brand" }),
  ]);
  const catalog = {
    categories: getIbsherApiList(categoryResponse.data, ["category", "categories", "data"]),
    subcategories,
    collections,
    brands,
  };
  if (Object.values(catalog).some((entries) => !Array.isArray(entries) || !entries.length)) {
    throw new Error("IBSHER returned incomplete mapping data. No products were changed.");
  }
  return catalog;
}

async function getIbsherPendingVendorImports(options = {}) {
  requireAccountSectionAccess("integration-review");
  const credentials = await resolveIbsherVendorAccount(options);
  return ibsherIntegrationReview.collectPendingImports({
    listImports: (pageOptions) => fetchIbsherVendorImportPage(pageOptions, credentials),
    shouldCancel: () => false,
  });
}

async function exportIbsherIntegrationReview(options = {}) {
  requireAccountSectionAccess("integration-review");
  const credentials = await resolveIbsherVendorAccount(options);
  const summaries = await ibsherIntegrationReview.collectImports({
    listImports: (pageOptions) => fetchIbsherVendorImportPage(pageOptions, credentials),
    shouldCancel: () => false,
  });
  if (!summaries.length) throw new Error("There are no IBSHER Integration products to export.");

  // The list endpoint contains summaries only. Vendor Data (including Vendor
  // Brand, vendor price and vendor stock) is available on the View/Review
  // detail endpoint, so hydrate every row before building the workbook.
  const products = new Array(summaries.length);
  let nextIndex = 0;
  const worker = async () => {
    while (nextIndex < summaries.length) {
      const index = nextIndex++;
      const details = await loadIbsherVendorImport(summaries[index]._id, credentials);
      products[index] = { ...summaries[index], ...details };
    }
  };
  await Promise.all(Array.from({ length: Math.min(5, summaries.length) }, worker));

  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Download all IBSHER Integration data as Excel",
    defaultPath: `ibsher-integration-all-${new Date().toISOString().slice(0, 10)}.xlsx`,
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (result.canceled || !result.filePath) return { canceled: true };

  const xlsx = loadXlsx();
  const { workbook, exportData } = ibsherIntegrationReview.createIntegrationReviewWorkbook(xlsx, products);
  xlsx.writeFile(workbook, result.filePath);
  return { canceled: false, filePath: result.filePath, exportedRows: exportData.rows.length };
}

async function processIbsherPendingVendorImports(options = {}) {
  requireAccountSectionAccess("integration-review");
  if (activeIntegrationReview) throw new Error("An Integration Review batch is already running.");
  const run = { canceled: false, userValue: authSession?.user?.value };
  activeIntegrationReview = run;
  try {
    const credentials = await resolveIbsherVendorAccount(options);
    const vendorAccountValue = getRequestedVendorValue(options);
    return await ibsherIntegrationReview.processPendingImports({}, {
      listImports: (pageOptions) => fetchIbsherVendorImportPage(pageOptions, credentials),
      loadImport: (importId) => loadIbsherVendorImport(importId, credentials),
      loadCatalog: () => loadIbsherIntegrationCatalog(credentials),
      saveMapping: async (importId, mapping) => (await requestIbsherAdmin({
        method: "PATCH", url: `/vendor-product-imports/${encodeURIComponent(importId)}`, data: mapping,
      }, credentials)).data,
      approve: async (importId, productData) => (await requestIbsherAdmin({
        method: "POST", url: `/vendor-product-imports/${encodeURIComponent(importId)}/approve`, data: productData,
      }, credentials)).data,
      shouldCancel: () => run.canceled || authSession?.user?.value !== run.userValue,
      onProgress: (progress) => sendToRenderer("ibsher-integration-review:progress", { ...progress, vendorAccountValue }),
    });
  } finally {
    activeIntegrationReview = null;
  }
}

async function fetchIbsherPublishingProductPage(options = {}, credentials) {
  const page = Math.max(1, Number.parseInt(options.page, 10) || 1);
  const limit = Math.max(1, Math.min(100, Number.parseInt(options.limit, 10) || 20));
  const status = options.status && options.status !== "all" ? ibsherProductPublishing.validateProductStatus(options.status) : "";
  const search = cleanCell(options.search).slice(0, 160);
  const response = await requestIbsherAdmin({
    method: "GET", url: "/product", params: { page, limit, ...(search ? { search } : {}), ...(status ? { status } : {}) },
  }, credentials);
  if (response.data?.success === false || !Array.isArray(response.data?.products)) {
    throw new Error(cleanCell(response.data?.message) || "IBSHER returned an invalid product list.");
  }
  return {
    products: response.data.products.map((product) => ({ ...addIbsherProductDisplayUrls(product), status: cleanCell(product.status) })),
    pagination: normalizeIbsherPagination(response.data.pagination, { page, limit }),
  };
}

async function fetchIbsherPublishingProducts(options = {}, credentials) {
  const filters = ibsherProductPublishing.normalizePublishingFilters(options.filters);
  if (!filters.brandId && !filters.collectionIds.length) {
    return fetchIbsherPublishingProductPage(options, credentials);
  }
  const requestedPage = Math.max(1, Number.parseInt(options.page, 10) || 1);
  const requestedLimit = Math.max(1, Math.min(100, Number.parseInt(options.limit, 10) || 20));
  const matched = [];
  const seen = new Set();
  for (let page = 1; page <= 10000; page += 1) {
    const result = await fetchIbsherPublishingProductPage({ ...options, filters: undefined, page, limit: 100 }, credentials);
    for (const product of result.products) {
      if (seen.has(product._id)) continue;
      seen.add(product._id);
      if (ibsherProductPublishing.matchesPublishingFilters(product, filters)) matched.push(product);
    }
    if (page >= result.pagination.totalPages) break;
    if (!result.products.length) throw new Error("Could not load the complete filtered IBSHER product list.");
  }
  const totalProducts = matched.length;
  const totalPages = Math.max(1, Math.ceil(totalProducts / requestedLimit));
  const currentPage = Math.min(requestedPage, totalPages);
  const start = (currentPage - 1) * requestedLimit;
  return {
    products: matched.slice(start, start + requestedLimit),
    pagination: { currentPage, limit: requestedLimit, totalProducts, totalPages, hasNextPage: currentPage < totalPages },
  };
}

async function getIbsherPublishingFilters(options = {}) {
  requireAccountSectionAccess("product-publishing");
  const credentials = await resolveIbsherVendorAccount(options);
  const products = new Map();
  for (let page = 1; page <= 10000; page += 1) {
    const result = await fetchIbsherPublishingProductPage({ page, limit: 100 }, credentials);
    for (const product of result.products) products.set(product._id, product);
    if (page >= result.pagination.totalPages) break;
    if (!result.products.length) throw new Error("Could not load the complete IBSHER product list for filters.");
  }
  const label = (entity) => cleanCell(entity?.name?.en || entity?.name?.ku || entity?.name?.ar || entity?.name || entity?.label);
  const id = (entity) => cleanCell(entity?._id || entity?.id || (typeof entity === "string" ? entity : ""));
  const brands = new Map();
  const categories = new Map();
  const subCategories = new Map();
  const collections = new Map();
  for (const product of products.values()) {
    const brand = product.brand;
    const collection = product.collectionName;
    const subCategory = product.subCategory || collection?.subCategory || collection?.subcategory;
    const category = product.category || subCategory?.category || collection?.category;
    const brandId = id(brand);
    const categoryId = id(category);
    const subCategoryId = id(subCategory);
    const collectionId = id(collection);
    if (brandId && label(brand)) brands.set(brandId, { id: brandId, name: label(brand) });
    if (categoryId && label(category)) categories.set(categoryId, { id: categoryId, name: label(category) });
    if (subCategoryId && label(subCategory)) subCategories.set(subCategoryId, {
      id: subCategoryId, name: label(subCategory), categoryId,
    });
    if (collectionId && label(collection)) collections.set(collectionId, {
      id: collectionId, name: label(collection), categoryId, subCategoryId,
    });
  }
  const sorted = (values) => [...values.values()].sort((left, right) => left.name.localeCompare(right.name));
  return { brands: sorted(brands), categories: sorted(categories), subCategories: sorted(subCategories), collections: sorted(collections) };
}

async function getIbsherPublishingProducts(options = {}) {
  requireAccountSectionAccess("product-publishing");
  const credentials = await resolveIbsherVendorAccount(options);
  return fetchIbsherPublishingProducts(options, credentials);
}

async function loadIbsherPublishingProduct(product, credentials) {
  if (!cleanCell(product.itemCode)) throw new Error("The selected product is missing its itemcode. Reload products first.");
  for (let page = 1; page <= 1000; page += 1) {
    const result = await fetchIbsherPublishingProducts({ search: product.itemCode, page, limit: 100 }, credentials);
    const match = result.products.find((candidate) => candidate._id === product._id);
    if (match) return match;
    if (page >= result.pagination.totalPages) return null;
  }
  throw new Error("Could not verify this product in the selected IBSHER account.");
}

async function changeIbsherProductStatuses(options = {}, publishAll = false) {
  requireAccountSectionAccess("product-publishing");
  if (activeProductPublishing) throw new Error("A product status update is already running.");
  const run = { canceled: false, userValue: authSession?.user?.value };
  activeProductPublishing = run;
  try {
    const sourceCredentials = await resolveIbsherVendorAccount(options);
    const accounts = await listVendorAccounts();
    const karimAccount = ibsherProductPublishing.selectKarimPublishingAccount(accounts);
    const publishingCredentials = await resolveIbsherVendorAccount({ vendorAccountValue: karimAccount.value });
    const vendorAccountValue = getRequestedVendorValue(options);
    const services = {
      // Product scope always comes from the Vendor selected in Product Publishing.
      listProducts: (pageOptions) => fetchIbsherPublishingProducts(pageOptions, sourceCredentials),
      // Confirm the ID belongs to the selected Vendor, then read it through karim
      // before and after every status write.
      loadProduct: async (product) => {
        const sourceProduct = await loadIbsherPublishingProduct(product, sourceCredentials);
        if (!sourceProduct) return null;
        return loadIbsherPublishingProduct(product, publishingCredentials);
      },
      updateStatus: async (request) => (await requestIbsherAdmin(request, publishingCredentials)).data,
      shouldCancel: () => run.canceled || authSession?.user?.value !== run.userValue,
      onProgress: (progress) => sendToRenderer("ibsher-publishing:progress", { ...progress, vendorAccountValue }),
    };
    return await (publishAll ? ibsherProductPublishing.publishAllProducts(services, options.filters) : ibsherProductPublishing.changeProductStatuses(options, services));
  } finally {
    activeProductPublishing = null;
  }
}

async function selectIbsherBrandWorkbook() {
  const selection = await dialog.showOpenDialog(mainWindow, {
    title: "Choose brand names and image links Excel",
    properties: ["openFile"],
    filters: [{ name: "Spreadsheets", extensions: ["xlsx", "xls", "csv"] }],
  });
  if (selection.canceled || !selection.filePaths[0]) return { canceled: true };
  return parseIbsherBrandWorkbook({ filePath: selection.filePaths[0] });
}

async function listIbsherBrandsForImport(credentials) {
  return fetchIbsherAdminListPages("/brand", credentials, ["brands", "brand", "data"], {
    label: "brand",
    limit: 500,
    extractList: (payload) => ibsherBrandImport.getBrandList(payload),
    getItemKey: (entry) => cleanCell(entry._id || entry.id) || ibsherBrandImport.brandKey(entry.name),
    incompleteMessage: "Could not load the complete IBSHER brand list. No brands were imported.",
    tooLargeMessage: "The IBSHER brand list is too large to load completely.",
  });
}

async function runIbsherBrandImport(options = {}) {
  requireAuthenticatedUser();
  if (activeBrandImport) throw new Error("A brand import is already running.");
  const run = { canceled: false };
  activeBrandImport = run;
  try {
    const credentials = await resolveIbsherVendorAccount(options);
    const vendorAccountValue = getRequestedVendorValue(options);
    return await ibsherBrandImport.importBrandRows(options.rows, {
      listBrands: () => listIbsherBrandsForImport(credentials),
      downloadImage: async (imageUrl) => {
        const response = await axios.get(imageUrl, {
          responseType: "arraybuffer",
          timeout: 30000,
          maxContentLength: ibsherBrandImport.MAX_BRAND_IMAGE_BYTES,
          maxRedirects: 5,
          headers: { Accept: "image/*" },
          beforeRedirect: (redirectOptions) => {
            if (!ibsherPageImages.parsePublicPageUrl(`${redirectOptions.protocol}//${redirectOptions.hostname}${redirectOptions.path || "/"}`)) {
              throw new Error("The image link redirects to an invalid public URL.");
            }
          },
        });
        return Buffer.from(response.data);
      },
      createBrand: async (row, buffer) => {
        const response = await requestIbsherAdmin(() => {
          const form = ibsherBrandImport.createBrandForm(row, buffer);
          return { method: "POST", url: "/brand", data: form, headers: form.getHeaders() };
        }, credentials);
        return response.data;
      },
      shouldCancel: () => run.canceled,
      onProgress: (progress) => sendToRenderer("ibsher-brands:progress", { ...progress, vendorAccountValue }),
    });
  } finally {
    activeBrandImport = null;
  }
}

async function importIbsherPageImageExcel() {
  requireAuthenticatedUser();
  const selection = await dialog.showOpenDialog(mainWindow, {
    title: "Choose web page links Excel",
    properties: ["openFile"],
    filters: [
      { name: "Excel Workbooks", extensions: ["xlsx", "xls", "csv"] },
    ],
  });
  if (selection.canceled || !selection.filePaths[0]) {
    return { canceled: true, rows: [], validCount: 0, invalidCount: 0 };
  }

  const filePath = selection.filePaths[0];
  const xlsx = loadXlsx();
  const workbook = xlsx.readFile(filePath, { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("The selected Excel workbook has no sheets.");
  const matrix = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: false,
  });
  const parsed = ibsherPageImages.parsePageWorkbookMatrix(matrix);
  if (!parsed.rows.length) throw new Error("The selected Excel sheet has no data rows.");

  return {
    canceled: false,
    filePath,
    sheetName,
    ...parsed,
  };
}

async function fetchIbsherProductPageImages(productId) {
  const normalizedId = String(productId || "").trim().toLowerCase();
  if (!/^[a-f\d]{24}$/.test(normalizedId)) {
    throw new Error("The product page link has an invalid product ID.");
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const endpoint = `https://api.ibsher.com/api/v1/client/product/${normalizedId}`;
    const response = await fetchWithRetry(endpoint, {
      headers: {
        Accept: "application/json",
        "Accept-Language": "en-US,en;q=0.9",
      },
      redirect: "follow",
      signal: controller.signal,
    }, "IBSHER product lookup", 2);
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || payload?.success === false) {
      throw new Error(payload?.message || `IBSHER product lookup returned HTTP ${response.status}.`);
    }
    const result = ibsherPageImages.extractIbsherProductImages(payload);
    if (!result.images.length) throw new Error("No product images were found for this page.");
    return result;
  } finally {
    clearTimeout(timeout);
  }
}

async function renderPageHtmlForImageExtraction(pageUrl) {
  let extractionWindow;
  let timeout;
  try {
    extractionWindow = new BrowserWindow({
      show: false,
      width: 1024,
      height: 768,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: true,
        javascript: true,
        partition: "page-image-extractor",
      },
    });
    extractionWindow.setMenu(null);
    extractionWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
    extractionWindow.webContents.session.setPermissionRequestHandler((webContents, permission, callback) => {
      callback(false);
    });

    await new Promise((resolve, reject) => {
      let settled = false;
      const finish = (error) => {
        if (settled) return;
        settled = true;
        clearTimeout(timeout);
        if (error) reject(error);
        else resolve();
      };
      timeout = setTimeout(() => finish(new Error("The web page took too long to load.")), 20000);
      extractionWindow.webContents.once("did-finish-load", () => finish());
      extractionWindow.webContents.once("did-fail-load", (event, code, description, validatedUrl, isMainFrame) => {
        if (isMainFrame !== false) finish(new Error(description || `The web page failed to load (${code}).`));
      });
      extractionWindow.loadURL(pageUrl, {
        userAgent: PRODUCT_SEARCH_HEADERS["User-Agent"],
      }).catch(finish);
    });

    await new Promise((resolve) => setTimeout(resolve, 1800));
    const rendered = await extractionWindow.webContents.executeJavaScript(`(() => {
      const goodContext = /product|gallery|carousel|media|photo|picture|zoom|main|primary|pdp|packshot/i;
      const badContext = /logo|icon|avatar|badge|flag|payment|rating|sprite|banner|advert|recommend|related|footer|header|nav/i;
      const images = Array.from(document.images).map((image, index) => {
        const url = image.currentSrc || image.src || '';
        const rectangle = image.getBoundingClientRect();
        const context = [
          image.alt,
          image.id,
          image.className,
          image.parentElement?.id,
          image.parentElement?.className,
          image.closest('figure, picture, section, article, main, div')?.className,
        ].join(' ');
        const naturalArea = Number(image.naturalWidth || 0) * Number(image.naturalHeight || 0);
        const visibleArea = Math.max(0, rectangle.width) * Math.max(0, rectangle.height);
        let score = Math.min(4000000, Math.max(naturalArea, visibleArea));
        if (goodContext.test(context)) score += 10000000;
        if (badContext.test(context)) score -= 20000000;
        if (image.fetchPriority === 'high') score += 500000;
        if (rectangle.top >= 0 && rectangle.top < window.innerHeight * 1.5) score += 100000;
        return { url, score, index, width: image.naturalWidth, height: image.naturalHeight };
      }).filter((image) => /^https?:\/\//i.test(image.url)
        && image.width >= 120
        && image.height >= 120
        && image.score > 0)
        .sort((left, right) => right.score - left.score || left.index - right.index)
        .map((image) => image.url);
      return {
        url: location.href,
        html: document.documentElement ? document.documentElement.outerHTML.slice(0, 8388608) : '',
        images: Array.from(new Set(images)).slice(0, 250),
      };
    })()`, true);
    const finalUrl = ibsherPageImages.parsePublicPageUrl(rendered?.url) || pageUrl;
    return {
      url: finalUrl,
      html: String(rendered?.html || ""),
      images: Array.isArray(rendered?.images) ? rendered.images : [],
    };
  } finally {
    clearTimeout(timeout);
    if (extractionWindow && !extractionWindow.isDestroyed()) extractionWindow.destroy();
  }
}

async function extractPrimaryImageFromAnyPage(pageUrl) {
  const normalizedPageUrl = ibsherPageImages.parsePublicPageUrl(pageUrl);
  if (!normalizedPageUrl) throw new Error("This is not a valid public web page link.");
  if (/\.(?:jpe?g|png|webp|avif)(?:[?#].*)?$/i.test(normalizedPageUrl)) {
    return { imageUrl: normalizedPageUrl, images: [normalizedPageUrl] };
  }

  const ibsherPage = ibsherPageImages.parseIbsherProductPageUrl(normalizedPageUrl);
  if (ibsherPage) return fetchIbsherProductPageImages(ibsherPage.productId);

  const staticPage = await fetchProductSearchPage(normalizedPageUrl, "Web page");
  if (staticPage) {
    const adapter = productImageSearch.detectProductPageAdapter(staticPage.url, staticPage.html);
    const imageUrls = productImageSearch.extractProductImages(
      staticPage.html,
      staticPage.url,
      {},
      { maxImages: 250 },
    );
    if (imageUrls[0]) return { imageUrl: imageUrls[0], images: imageUrls };

    const dataUrls = productImageSearch.buildProductImageDataUrls(staticPage.url, adapter);
    for (const dataUrl of dataUrls) {
      const dataPage = await fetchProductSearchPage(dataUrl, "Product image data");
      const dataImages = dataPage
        ? productImageSearch.extractProductImagesFromData(
          dataPage.html,
          staticPage.url,
          adapter,
          { maxImages: 250 },
        )
        : [];
      if (dataImages[0]) return { imageUrl: dataImages[0], images: dataImages };
    }
  }

  const renderedPage = await renderPageHtmlForImageExtraction(normalizedPageUrl);
  const renderedImages = productImageSearch.extractProductImages(
    renderedPage.html,
    renderedPage.url,
    {},
    { maxImages: 250 },
  );
  const images = [...new Map([...renderedImages, ...renderedPage.images].map((imageUrl) => (
    [productImageSearch.imageCanonicalKey(imageUrl), imageUrl]
  ))).values()].slice(0, 250);
  if (!images[0]) throw new Error("No product image was found on this web page.");
  return { imageUrl: images[0], images };
}

async function extractIbsherPageImageLinks(options = {}) {
  requireAuthenticatedUser();
  const sourceRows = Array.isArray(options.rows) ? options.rows : [];
  if (!sourceRows.length) throw new Error("Import an Excel file before extracting image links.");

  const readyIndexes = sourceRows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => row?.status === "ready" && row?.pageUrl);
  if (!readyIndexes.length) throw new Error("There are no valid product page links ready to process.");

  let processed = 0;
  const processedRows = await mapWithConcurrency(readyIndexes, 3, async ({ row, index }) => {
    let nextRow;
    try {
      const result = await extractPrimaryImageFromAnyPage(row.pageUrl);
      nextRow = {
        ...row,
        ...result,
        status: "success",
        error: "",
      };
    } catch (error) {
      nextRow = {
        ...row,
        imageUrl: "",
        images: [],
        status: "error",
        error: error.message || "Could not extract product image links.",
      };
    }
    processed += 1;
    sendToRenderer("ibsher-page-images:progress", {
      processed,
      total: readyIndexes.length,
      rowNumber: row.rowNumber,
      succeeded: nextRow.status === "success",
    });
    return { index, row: nextRow };
  });

  const rows = sourceRows.map((row) => ({ ...row }));
  processedRows.forEach((entry) => {
    rows[entry.index] = entry.row;
  });
  return {
    rows,
    processed,
    total: readyIndexes.length,
    succeeded: rows.filter((row) => row.status === "success").length,
    failed: rows.filter((row) => row.status === "error").length,
    imageCount: rows.reduce((sum, row) => sum + (Array.isArray(row.images) ? row.images.length : 0), 0),
  };
}

async function exportIbsherPageImageExcel(options = {}) {
  requireAuthenticatedUser();
  const exportData = ibsherPageImages.buildIbsherPageImageExcelData(options.rows);
  const sourceFilePath = String(options.sourceFilePath || "").trim();
  const sourceName = sourceFilePath
    ? path.basename(sourceFilePath, path.extname(sourceFilePath))
    : "page-images";
  const defaultPath = sourceFilePath
    ? path.join(path.dirname(sourceFilePath), `${sourceName}-with-images.xlsx`)
    : `${sourceName}-${Date.now()}.xlsx`;
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Save image links Excel",
    defaultPath,
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (result.canceled || !result.filePath) return { canceled: true };

  const xlsx = loadXlsx();
  const worksheet = xlsx.utils.json_to_sheet(exportData.excelRows, { header: exportData.columns });
  worksheet["!cols"] = exportData.columns.map(() => ({ wch: 72 }));
  exportData.excelRows.forEach((row, rowIndex) => {
    exportData.imageColumns.forEach((column) => {
      const value = row[column];
      if (!/^https?:\/\//i.test(String(value || ""))) return;
      const columnIndex = exportData.columns.indexOf(column);
      const address = xlsx.utils.encode_cell({ r: rowIndex + 1, c: columnIndex });
      if (worksheet[address]) worksheet[address].l = { Target: value };
    });
  });
  worksheet["!autofilter"] = {
    ref: xlsx.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: exportData.excelRows.length, c: exportData.columns.length - 1 },
    }),
  };
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Product Images");
  xlsx.writeFile(workbook, result.filePath);
  return { canceled: false, filePath: result.filePath, exported: exportData.excelRows.length };
}

function getPostimagesMimeType(filePath) {
  return {
    ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
    ".webp": "image/webp", ".gif": "image/gif", ".bmp": "image/bmp",
    ".avif": "image/avif", ".heic": "image/heic", ".heif": "image/heif",
  }[path.extname(filePath).toLowerCase()] || "application/octet-stream";
}

async function uploadFileToPostimages(filePath, settings) {
  if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
    throw new Error(`Image not found: ${filePath}`);
  }
  const originalBuffer = fs.readFileSync(filePath);
  if (!originalBuffer.length) throw new Error(`Image is empty: ${path.basename(filePath)}`);
  if (originalBuffer.length > 32 * 1024 * 1024) {
    throw new Error(`${path.basename(filePath)} exceeds the 32 MB Postimages free-account limit.`);
  }

  const originalMimeType = getPostimagesMimeType(filePath);
  const prepared = await prepareAiProductImageBuffer(originalBuffer, originalMimeType, filePath);
  const converted = prepared.mimeType === "image/jpeg" && prepared.mimeType !== originalMimeType;
  const filename = converted ? `${path.parse(filePath).name}.jpg` : path.basename(filePath);
  const form = new URLSearchParams({
    key: settings.apiKey,
    gallery: settings.gallery || "",
    o: "2b819584285c102318568238c7d4a4c7",
    m: "59c2ad4b46b0c1e12d5703302bff0120",
    version: "1.0.1",
    portable: "1",
    name: path.parse(filename).name,
    type: path.extname(filename).slice(1).toLowerCase(),
    image: prepared.buffer.toString("base64"),
  });

  let response;
  try {
    response = await axios.post("https://api.postimage.org/1/upload", form.toString(), {
      headers: { "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8" },
      timeout: 120000,
      maxBodyLength: Infinity,
      maxContentLength: Infinity,
      validateStatus: () => true,
    });
  } catch (error) {
    throw new Error(`Postimages upload failed for ${path.basename(filePath)}: ${error.message}`);
  }
  if (response.status < 200 || response.status >= 300) {
    const message = response.data?.error || response.data?.message || `HTTP ${response.status}`;
    throw new Error(`Postimages rejected ${path.basename(filePath)}: ${message}`);
  }
  try {
    return { path: filePath, name: path.basename(filePath), ...extractPostimagesUploadResult(response.data) };
  } catch (error) {
    throw new Error(`${path.basename(filePath)}: ${error.message}`);
  }
}

async function selectPostimagesImages() {
  requireAccountSectionAccess("postimages-uploader");
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Choose images to upload to Postimages",
    properties: ["openFile", "multiSelections"],
    filters: [{ name: "Images", extensions: [...POSTIMAGES_IMAGE_EXTENSIONS].map((value) => value.slice(1)) }],
  });
  if (result.canceled) return { canceled: true, images: [] };
  return {
    canceled: false,
    images: result.filePaths.map((filePath) => ({
      path: filePath,
      name: path.basename(filePath),
      size: fs.statSync(filePath).size,
    })),
  };
}

function findPostimagesExcelEntries(workbook, workbookPath) {
  const entries = [];
  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    for (const [cellAddress, cell] of Object.entries(worksheet)) {
      if (cellAddress.startsWith("!") || cell?.v === undefined || cell?.v === null) continue;
      const imagePath = resolvePostimagesLocalPath(cell.v, workbookPath);
      if (!imagePath) continue;
      entries.push({ sheetName, cellAddress, path: imagePath, name: path.basename(imagePath), source: "path" });
    }
  }
  return entries;
}

async function selectPostimagesExcel() {
  requireAccountSectionAccess("postimages-uploader");
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Choose an Excel file containing image paths or embedded images",
    properties: ["openFile"],
    filters: [{ name: "Excel Workbook", extensions: ["xlsx", "xls", "csv"] }],
  });
  if (result.canceled || !result.filePaths[0]) return { canceled: true, entries: [] };
  const filePath = result.filePaths[0];
  const workbook = loadXlsx().readFile(filePath);
  const temporaryImageDirectory = path.extname(filePath).toLowerCase() === ".xlsx"
    ? fs.mkdtempSync(path.join(app.getPath("temp"), "postimages-excel-preview-"))
    : "";
  try {
    const embeddedEntries = temporaryImageDirectory
      ? await extractEmbeddedImagesFromWorkbook(filePath, temporaryImageDirectory)
      : [];
    const entries = [...findPostimagesExcelEntries(workbook, filePath), ...embeddedEntries];
    return {
      canceled: false,
      filePath,
      entries: entries.map((entry) => ({
        sheetName: entry.sheetName,
        cellAddress: entry.cellAddress,
        name: entry.name,
        source: entry.source,
      })),
    };
  } finally {
    if (temporaryImageDirectory) fs.rmSync(temporaryImageDirectory, { recursive: true, force: true });
  }
}

async function uploadPostimagesImages(options = {}) {
  requireAccountSectionAccess("postimages-uploader");
  const settings = readPostimagesSettings();
  if (!settings.apiKey) throw new Error("Save your official Postimages API key before uploading.");
  const imagePaths = [...new Set((options.imagePaths || []).map((value) => String(value || "").trim()).filter(Boolean))];
  if (!imagePaths.length) throw new Error("Choose at least one image.");
  if (imagePaths.length > 1000) throw new Error("Postimages accepts no more than 1,000 images per batch.");
  const results = [];
  for (let index = 0; index < imagePaths.length; index += 1) {
    sendToRenderer("postimages:progress", { current: index, total: imagePaths.length, name: path.basename(imagePaths[index]) });
    try {
      results.push({ ok: true, ...(await uploadFileToPostimages(imagePaths[index], settings)) });
    } catch (error) {
      results.push({ ok: false, path: imagePaths[index], name: path.basename(imagePaths[index]), error: error.message });
    }
  }
  sendToRenderer("postimages:progress", { current: imagePaths.length, total: imagePaths.length, complete: true });
  return { results };
}

async function uploadPostimagesExcel(options = {}) {
  requireAccountSectionAccess("postimages-uploader");
  const settings = readPostimagesSettings();
  if (!settings.apiKey) throw new Error("Save your official Postimages API key before uploading.");
  const workbookPath = String(options.filePath || "").trim();
  if (!workbookPath || !fs.existsSync(workbookPath)) throw new Error("Choose an Excel file first.");
  const xlsx = loadXlsx();
  const workbook = xlsx.readFile(workbookPath);
  const temporaryImageDirectory = path.extname(workbookPath).toLowerCase() === ".xlsx"
    ? fs.mkdtempSync(path.join(app.getPath("temp"), "postimages-excel-upload-"))
    : "";

  try {
    const embeddedEntries = temporaryImageDirectory
      ? await extractEmbeddedImagesFromWorkbook(workbookPath, temporaryImageDirectory)
      : [];
    const entries = [...findPostimagesExcelEntries(workbook, workbookPath), ...embeddedEntries];
    if (!entries.length) throw new Error("No readable local image paths or embedded images were found in the Excel file.");

    const uniquePaths = [...new Set(entries.map((entry) => entry.path))];
    const uploadedByPath = new Map();
    const failures = [];
    for (let index = 0; index < uniquePaths.length; index += 1) {
      const imagePath = uniquePaths[index];
      sendToRenderer("postimages:progress", { current: index, total: uniquePaths.length, name: path.basename(imagePath) });
      try { uploadedByPath.set(imagePath, await uploadFileToPostimages(imagePath, settings)); }
      catch (error) { failures.push({ path: imagePath, name: path.basename(imagePath), error: error.message }); }
    }

    const failedEmbeddedPaths = new Set(
      entries.filter((entry) => entry.source === "embedded" && !uploadedByPath.has(entry.path)).map((entry) => entry.path)
    );
    if (failedEmbeddedPaths.size) {
      throw new Error(`${failedEmbeddedPaths.size} embedded Excel image(s) failed to upload. No output workbook was saved, so the original embedded images remain intact.`);
    }

    const urlsByCell = new Map();
    entries.forEach((entry) => {
      const uploaded = uploadedByPath.get(entry.path);
      if (!uploaded) return;
      const key = `${entry.sheetName}\u0000${entry.cellAddress}`;
      const urls = urlsByCell.get(key) || [];
      if (!urls.includes(uploaded.url)) urls.push(uploaded.url);
      urlsByCell.set(key, urls);
    });
    urlsByCell.forEach((urls, key) => {
      const [sheetName, cellAddress] = key.split("\u0000");
      const worksheet = workbook.Sheets[sheetName];
      const cell = worksheet[cellAddress] || (worksheet[cellAddress] = {});
      cell.v = urls.join("\n");
      cell.t = "s";
      delete cell.w;
      delete cell.h;
      const position = xlsx.utils.decode_cell(cellAddress);
      const range = worksheet["!ref"]
        ? xlsx.utils.decode_range(worksheet["!ref"])
        : { s: { ...position }, e: { ...position } };
      range.s.r = Math.min(range.s.r, position.r);
      range.s.c = Math.min(range.s.c, position.c);
      range.e.r = Math.max(range.e.r, position.r);
      range.e.c = Math.max(range.e.c, position.c);
      worksheet["!ref"] = xlsx.utils.encode_range(range);
    });

    const parsed = path.parse(workbookPath);
    const saveResult = await dialog.showSaveDialog(mainWindow, {
      title: "Save Excel with Postimages links",
      defaultPath: path.join(parsed.dir, `${parsed.name}-postimages.xlsx`),
      filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
    });
    if (saveResult.canceled || !saveResult.filePath) return { canceled: true, failures };
    xlsx.writeFile(workbook, saveResult.filePath);
    sendToRenderer("postimages:progress", { current: uniquePaths.length, total: uniquePaths.length, complete: true });
    return {
      canceled: false,
      filePath: saveResult.filePath,
      uploaded: uploadedByPath.size,
      replacedCells: urlsByCell.size,
      failures,
    };
  } finally {
    if (temporaryImageDirectory) fs.rmSync(temporaryImageDirectory, { recursive: true, force: true });
  }
}

async function selectAiProductImages() {
  requireAuthenticatedUser();
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Choose product images",
    properties: ["openFile", "multiSelections"],
    filters: [
      { name: "Product Images", extensions: ["jpg", "jpeg", "png", "webp", "heic", "heif"] },
    ],
  });

  if (result.canceled) return { canceled: true, images: [] };
  if (result.filePaths.length > 15) throw new Error("Choose no more than 15 images for one product.");

  const images = (await Promise.all(result.filePaths.map(createAiImageInfo))).filter(Boolean);
  const totalBytes = images.reduce((sum, image) => sum + image.size, 0);
  if (totalBytes > 18 * 1024 * 1024) {
    throw new Error("The selected images exceed 18 MB in total. Resize them before AI analysis.");
  }

  return { canceled: false, images };
}

function listAiProductFolderImages(folderPath, limit = 2000) {
  const imagePaths = [];
  const pendingFolders = [folderPath];
  while (pendingFolders.length && imagePaths.length <= limit) {
    const currentFolder = pendingFolders.shift();
    const entries = fs.readdirSync(currentFolder, { withFileTypes: true })
      .sort((left, right) => left.name.localeCompare(right.name, undefined, { numeric: true }));
    for (const entry of entries) {
      const entryPath = path.join(currentFolder, entry.name);
      if (entry.isDirectory()) pendingFolders.push(entryPath);
      else if (entry.isFile() && AI_PRODUCT_IMAGE_EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
        imagePaths.push(entryPath);
        if (imagePaths.length > limit) break;
      }
    }
  }
  if (imagePaths.length > limit) {
    throw new Error(`The selected folder contains more than ${limit} images. Choose a smaller product image folder.`);
  }
  return imagePaths;
}

async function selectAiProductImageFolder() {
  requireAuthenticatedUser();
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Choose a folder and group its images into products",
    properties: ["openDirectory"],
  });
  if (result.canceled || !result.filePaths[0]) return { canceled: true, images: [] };

  const folderPath = result.filePaths[0];
  const imagePaths = listAiProductFolderImages(folderPath);
  const images = imagePaths.map((imagePath) => {
    const stats = fs.statSync(imagePath);
    return {
      path: imagePath,
      name: path.basename(imagePath),
      relativeName: path.relative(folderPath, imagePath),
      size: stats.size,
      mimeType: getImageMimeType(imagePath),
      previewUrl: "",
    };
  });

  return { canceled: false, folderPath, images };
}

async function loadAiProductImagePreviews(options = {}) {
  requireAuthenticatedUser();
  const imagePaths = Array.isArray(options.imagePaths)
    ? [...new Set(options.imagePaths.map((value) => String(value || "").trim()).filter(Boolean))]
    : [];
  if (imagePaths.length > 24) throw new Error("No more than 24 image previews can be loaded at once.");

  const images = [];
  const skipped = [];
  for (let start = 0; start < imagePaths.length; start += 4) {
    const preparedBatch = await Promise.all(imagePaths.slice(start, start + 4).map(async (imagePath) => {
      try {
        const image = await createAiImageInfo(imagePath, { thumbnail: true });
        if (!image) throw new Error("The image is unavailable or unsupported.");
        return image;
      } catch (error) {
        skipped.push({ path: imagePath, name: path.basename(imagePath), reason: error.message });
        return null;
      }
    }));
    images.push(...preparedBatch.filter(Boolean));
  }
  return { images, skipped };
}

async function loadAiProductImagePreview(options = {}) {
  requireAuthenticatedUser();
  const imagePath = String(options.imagePath || "").trim();
  if (!imagePath) throw new Error("Choose an image to preview.");
  const image = await createAiImageInfo(imagePath);
  if (!image) throw new Error("The selected image is unavailable or unsupported.");
  return image;
}

async function pasteAiProductImage() {
  requireAuthenticatedUser();
  const clipboardImage = clipboard.readImage();
  if (clipboardImage.isEmpty()) {
    throw new Error("No image was found in the clipboard. Copy an image first, then try Paste Image again.");
  }
  const buffer = clipboardImage.toPNG();
  if (!buffer.length) throw new Error("The clipboard image could not be read.");
  const imageDirectory = path.join(app.getPath("userData"), "ai-product-images");
  fs.mkdirSync(imageDirectory, { recursive: true });
  const filePath = path.join(imageDirectory, `pasted-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.png`);
  fs.writeFileSync(filePath, buffer);
  const image = await createAiImageInfo(filePath);
  if (!image) throw new Error("The pasted image could not be prepared.");
  return { image };
}

async function importAiProductsFromExcel() {
  requireAuthenticatedUser();
  const selection = await dialog.showOpenDialog(mainWindow, {
    title: "Import AI product batch",
    properties: ["openFile"],
    filters: [
      { name: "Excel Workbook", extensions: ["xlsx", "xls", "csv"] },
    ],
  });
  if (selection.canceled || !selection.filePaths[0]) {
    return { canceled: true, products: [], skipped: [] };
  }

  const workbookPath = selection.filePaths[0];
  const xlsx = loadXlsx();
  const workbook = xlsx.readFile(workbookPath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  const imageColumnIndexes = Object.entries(worksheet)
    .filter(([cellAddress, cell]) => {
      if (cellAddress.startsWith("!")) return false;
      const position = xlsx.utils.decode_cell(cellAddress);
      const header = String(cell?.v || "").trim().toLowerCase();
      return position.r === 0 && (header === "image" || header === "images" || /^image-\d+$/.test(header));
    })
    .map(([cellAddress]) => xlsx.utils.decode_cell(cellAddress).c);
  const embeddedImagesByRow = await extractEmbeddedImagesFromFirstWorksheet(
    workbookPath,
    path.join(app.getPath("userData"), "ai-product-images", "excel"),
    imageColumnIndexes,
    workbook.SheetNames[0]
  );
  const products = [];
  const skipped = [];

  for (const [index, row] of rows.entries()) {
    const normalizedEntries = Object.entries(row).reduce((result, [key, value]) => {
      result[String(key).trim().toLowerCase()] = value;
      return result;
    }, {});
    const imageValues = getRowImageValues(normalizedEntries);
    const resolvedImageSources = resolveExcelImageSources(imageValues, {
      workbookFile: workbookPath,
      imageExtensions: AI_PRODUCT_IMAGE_EXTENSIONS,
    });
    const imageSources = [
      ...resolvedImageSources.sources,
      ...(embeddedImagesByRow.get(index + 2) || []),
    ];
    let images;
    try {
      images = (await Promise.all(imageSources.map((imageSource) => {
        if (isRemoteImageSource(imageSource)) return createRemoteAiImageInfo(imageSource);
        return createAiImageInfo(imageSource);
      }))).filter(Boolean);
    } catch (error) {
      skipped.push({ rowNumber: index + 2, reason: error.message });
      continue;
    }

    if (imageValues.length && !images.length) {
      skipped.push({
        rowNumber: index + 2,
        reason: resolvedImageSources.errors[0] || "No product images could be read from the Excel path.",
      });
      continue;
    }

    if (images.length > 15) {
      skipped.push({ rowNumber: index + 2, reason: "More than 15 images were provided." });
      continue;
    }
    if (images.reduce((sum, image) => sum + image.size, 0) > 18 * 1024 * 1024) {
      skipped.push({ rowNumber: index + 2, reason: "Images exceed 18 MB in total." });
      continue;
    }

    const importedRow = {
      itemcode: cleanCell(normalizedEntries.itemcode || normalizedEntries.item_code),
      en: cleanCell(normalizedEntries.en || normalizedEntries.in),
      ar: cleanCell(normalizedEntries.ar),
      ku: cleanCell(normalizedEntries.ku || normalizedEntries.on || normalizedEntries.to),
      "description-en": cleanCell(normalizedEntries["description-en"] || normalizedEntries["description-in"]),
      "description-ar": cleanCell(normalizedEntries["description-ar"]),
      "description-ku": cleanCell(
        normalizedEntries["description-ku"]
          || normalizedEntries["description-on"]
          || normalizedEntries["description-they"]
      ),
      collectionName: cleanCell(
        normalizedEntries.collectionname
          || normalizedEntries.collection_name
          || normalizedEntries.collection
      ),
      brand: cleanCell(normalizedEntries.brand),
      keyword: cleanCell(normalizedEntries.keyword),
      color: normalizeSingleColor(normalizedEntries.color),
      size: normalizeProductSize(normalizedEntries.size),
    };
    const mappedSourceHeaders = new Set([
      "itemcode", "item_code", "en", "in", "ar", "ku", "on", "to",
      "description-en", "description-in", "description-ar", "description-ku",
      "description-on", "description-they", "collectionname", "collection_name",
      "collection", "brand", "keyword", "color", "size", "images", "image",
    ]);
    Object.entries(row).forEach(([sourceHeader, value]) => {
      const normalizedHeader = String(sourceHeader).trim().toLowerCase();
      if (!mappedSourceHeaders.has(normalizedHeader) && !/^image-\d+$/.test(normalizedHeader)) {
        importedRow[sourceHeader] = value ?? "";
      }
    });
    images.forEach((image, imageIndex) => {
      importedRow[`image-${imageIndex + 1}`] = image.path;
    });

    if (!Object.values(importedRow).some(Boolean)) {
      skipped.push({ rowNumber: index + 2, reason: "The row is empty." });
      continue;
    }

    products.push({
      sourceRow: index + 2,
      images,
      row: importedRow,
      missingFields: [
        ...AI_EXCEL_COMPLETION_FIELDS.filter((field) => !cleanCell(importedRow[field])),
        ...(!images.length ? ["image-1"] : []),
      ],
    });
  }

  return {
    canceled: false,
    filePath: workbookPath,
    products,
    skipped,
  };
}

function extractGeminiJson(response) {
  const text = response?.candidates?.[0]?.content?.parts
    ?.map((part) => part.text || "")
    .join("")
    .trim();
  if (!text) throw new Error("Gemini returned an empty product result.");

  try {
    return JSON.parse(text.replace(/^```json\s*|\s*```$/g, ""));
  } catch {
    throw new Error("Gemini returned product data in an invalid format. Try again with clearer images.");
  }
}

async function analyzeProductImages(imagePaths, hints, settings, existingRow = {}, ibsherCatalog = {}) {
  let totalImageBytes = 0;
  let payloadImageBytes = 0;
  const imageParts = [];
  const pendingImageSources = [...imagePaths];
  const seenImageSources = new Set();
  while (pendingImageSources.length) {
    const imageSource = pendingImageSources.shift();
    const sourceKey = String(imageSource || "").trim();
    if (!sourceKey || seenImageSources.has(sourceKey)) continue;
    seenImageSources.add(sourceKey);
    let buffer;
    let mimeType;
    if (isRemoteImageSource(imageSource)) {
      try {
        const imageResponse = await fetchWithRetry(imageSource, {
          headers: { Accept: "image/*" },
        }, "Excel image download");
        if (!imageResponse.ok) {
          console.warn(`[ai-product] Skipping image URL with HTTP ${imageResponse.status}: ${imageSource}`);
          continue;
        }
        buffer = Buffer.from(await imageResponse.arrayBuffer());
        mimeType = String(imageResponse.headers.get("content-type") || "").split(";")[0];
        if (!mimeType.startsWith("image/")) {
          mimeType = getImageMimeType(new URL(imageSource).pathname);
        }
        if (!mimeType?.startsWith("image/") && isPostimgPageUrl(imageSource)) {
          const resolvedImageUrls = extractPostimgImageUrls(buffer.toString("utf8"), imageSource, 15)
            .filter((resolvedUrl) => !seenImageSources.has(resolvedUrl));
          if (!resolvedImageUrls.length) {
            throw new Error("The Postimages link did not contain any downloadable images.");
          }
          if (imageParts.length + pendingImageSources.length + resolvedImageUrls.length > 15) {
            throw new Error("A product can contain no more than 15 images, including images inside a gallery link.");
          }
          console.log(`[ai-product] Resolved Postimages page to ${resolvedImageUrls.length} direct image(s).`);
          pendingImageSources.unshift(...resolvedImageUrls);
          continue;
        }
      } catch (error) {
        if (/Postimages link|no more than 15 images/.test(error.message)) throw error;
        console.warn(`[ai-product] Skipping unavailable image URL: ${imageSource}`, error.message);
        continue;
      }
    } else {
      buffer = fs.readFileSync(imageSource);
      mimeType = getImageMimeType(imageSource);
    }
    if (!mimeType?.startsWith("image/")) throw new Error("An Excel image link did not return a supported image.");
    totalImageBytes += buffer.length;
    const compatibleImage = await prepareAiProductImageBuffer(buffer, mimeType, imageSource);
    buffer = compatibleImage.buffer;
    mimeType = compatibleImage.mimeType;
    const optimizedImage = optimizeAiProductImage(buffer, mimeType);
    buffer = optimizedImage.buffer;
    mimeType = optimizedImage.mimeType;
    payloadImageBytes += buffer.length;
    imageParts.push({
      inline_data: {
        mime_type: mimeType,
        data: buffer.toString("base64"),
      },
    });
  }
  if (totalImageBytes > 18 * 1024 * 1024) {
    throw new Error("The images for one Excel row exceed 18 MB in total.");
  }
  if (payloadImageBytes < totalImageBytes) {
    console.log(`[ai-product] Reduced Gemini image payload from ${totalImageBytes} to ${payloadImageBytes} bytes.`);
  }
  const catalog = sanitizeIbsherProductCatalog(ibsherCatalog);
  const collectionOptions = catalog.collections.map((collection) => ({
    collectionName: collection.name,
    subCategory: collection.subCategory,
    category: collection.category,
  }));
  const localizationInstructions = buildExistingLocalizationInstructions(existingRow);
  const localizedSchemaDescriptions = buildLocalizedSchemaDescriptions(existingRow);
  const hasCollectionCatalog = collectionOptions.length > 0;
  const prompt = [
    imageParts.length
      ? "Analyze all attached images as views of one retail product."
      : "Complete the missing catalog fields using the existing product row.",
    "Return accurate sales-catalog data in English, Arabic, and Central Kurdish (Sorani).",
    "Preserve the meaning of every non-empty existing field and only infer information needed for missing fields.",
    ...localizationInstructions,
    "Do not invent a brand, material, specification, or item code that is not visible or strongly supported.",
    hasCollectionCatalog
      ? "For collectionName, choose only one exact collectionName value from the IBSHER collection catalog below when the product is a clear fit."
      : "No collection catalog is available. Preserve an existing collectionName or operator hint; otherwise return an empty string.",
    hasCollectionCatalog
      ? "Use the category and subCategory only as matching context. Never return a category or subCategory as collectionName."
      : "Never invent a collectionName without a supplied existing value or operator hint.",
    hasCollectionCatalog
      ? "If none of the IBSHER collections is a suitable clear match, return an empty string for collectionName. Do not guess."
      : "Return an empty collectionName when no collection value was supplied.",
    "When Excel has no product name and a new name must be generated, keep the brand name only in the brand field and do not include it in en, ar, or ku.",
    "When translating an existing Excel product name, never add a brand that is absent from the source; if a brand is already part of the source text, keep that brand spelling unchanged.",
    "When Excel has no product description and a new description must be generated, write a useful natural description rather than image-position commentary.",
    "Return exactly one dominant product color as a single English color name. Never return multiple colors, a list, separators, or color alternatives.",
    "Return exactly one product size. Rules: 1. For clothing, return the exact label found (e.g., S, XL, 32). 2. For all other products, return ONLY 'Small', 'Medium', or 'Large'. 3. If no size is found, return 'Small'. No extra text.",
    "Return exactly 20 concise keyword concepts for each language: 20 English, 20 Arabic, and 20 Kurdish Sorani keywords (60 keywords total), with no empty values or duplicates.",
    "For every concept, keep keywordsEn, keywordsAr, and keywordsKu at the same array index as direct translations of each other.",
    `Optional operator hints: ${JSON.stringify(hints || {})}`,
    `Existing product row: ${JSON.stringify(existingRow || {})}`,
    ...(hasCollectionCatalog ? [`IBSHER collection catalog: ${JSON.stringify(collectionOptions)}`] : []),
  ].join("\n");
  const schema = {
    type: "OBJECT",
    properties: {
      itemcode: { type: "STRING" },
      en: { type: "STRING", description: localizedSchemaDescriptions.en },
      ar: { type: "STRING", description: localizedSchemaDescriptions.ar },
      ku: { type: "STRING", description: localizedSchemaDescriptions.ku },
      descriptionEn: { type: "STRING", description: localizedSchemaDescriptions.descriptionEn },
      descriptionAr: { type: "STRING", description: localizedSchemaDescriptions.descriptionAr },
      descriptionKu: { type: "STRING", description: localizedSchemaDescriptions.descriptionKu },
      collectionName: {
        type: "STRING",
        description: hasCollectionCatalog
          ? "One exact collectionName from the provided IBSHER collection catalog, or an empty string when no suitable match exists."
          : "The supplied existing collectionName or operator hint, otherwise an empty string.",
      },
      brand: { type: "STRING" },
      color: {
        type: "STRING",
        description: "Exactly one dominant product color in English; never a list or multiple colors.",
      },
      size: {
        type: "STRING",
        description: "Exactly one product size such as S, XL, 42, or One size.",
      },
      keywordsEn: { type: "ARRAY", items: { type: "STRING" }, minItems: 20, maxItems: 20 },
      keywordsAr: { type: "ARRAY", items: { type: "STRING" }, minItems: 20, maxItems: 20 },
      keywordsKu: { type: "ARRAY", items: { type: "STRING" }, minItems: 20, maxItems: 20 },
    },
    required: [
      "itemcode", "en", "ar", "ku", "descriptionEn", "descriptionAr", "descriptionKu",
      "collectionName", "brand", "color", "size", "keywordsEn", "keywordsAr", "keywordsKu",
    ],
  };
  const models = aiProductPerformance.getAiProductModelOrder();
  let lastError;
  for (const model of models) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
    const requestBody = JSON.stringify({
      contents: [{ parts: [...imageParts, { text: prompt }] }],
      generationConfig: aiProductPerformance.getAiProductGenerationConfig(model, schema),
    });
    try {
      const response = await fetchWithRetry(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-goog-api-key": settings.geminiApiKey,
        },
        body: requestBody,
      }, `Gemini model ${model}`, 4);
      const data = await response.json().catch(() => ({}));
      if (response.ok) {
        return {
          product: extractGeminiJson(data),
          api: {
            model,
            usage: normalizeGeminiUsage(data.usageMetadata),
          },
        };
      }
      const message = data?.error?.message || `Gemini model ${model} returned HTTP ${response.status}.`;
      if (response.status === 401 || response.status === 403) {
        const accessMessage = response.status === 401
          ? "The configured Google AI Studio API key is invalid. Update it in Profile settings."
          : "The configured Google AI Studio API key does not have access to Gemini 3.5 Flash-Lite. Check the key's API restrictions, project billing, and Gemini API access in Google AI Studio.";
        throw new Error(`${accessMessage} (${message})`);
      }
      lastError = new Error(message);
    } catch (error) {
      lastError = error;
      console.warn(`[ai-product] Gemini model ${model} failed:`, error.message);
    }
  }
  throw new Error(`Gemini is temporarily unavailable after trying ${models.length} models. ${lastError?.message || "Please retry the row."}`);
}

function normalizeKeywords(values) {
  return aiProductKeywords.normalizeKeywordList(values);
}

function formatMultilingualKeywords(keywordsEn, keywordsAr, keywordsKu) {
  return aiProductKeywords.formatMultilingualKeywords(keywordsEn, keywordsAr, keywordsKu);
}

function removeBrandFromProductName(value, brand) {
  const productName = cleanCell(value);
  const brandName = cleanCell(brand);
  if (!productName || !brandName) return productName;

  const escapedBrand = brandName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const brandPattern = new RegExp(
    `(^|[^\\p{L}\\p{N}])${escapedBrand}(?=$|[^\\p{L}\\p{N}])`,
    "giu"
  );

  return productName
    .replace(brandPattern, "$1")
    .replace(/^\s*['’]s\b/i, "")
    .replace(/\s{2,}/g, " ")
    .replace(/^[\s\-–—|:;,/]+|[\s\-–—|:;,/]+$/g, "")
    .trim();
}

async function generateAiProduct(options = {}) {
  requireAuthenticatedUser();
  const imagePaths = Array.isArray(options.imagePaths) ? options.imagePaths : [];
  if (!imagePaths.length) throw new Error("Choose at least one product image.");
  for (const imagePath of imagePaths) {
    if (!isRemoteImageSource(imagePath) && (!fs.existsSync(imagePath) || !getImageMimeType(imagePath))) {
      throw new Error("One of the selected images is unavailable or unsupported.");
    }
  }
  const settings = readAiSettings();
  if (!settings.geminiApiKey) throw new Error("Configure the Google AI Studio API key first.");
  const ibsherCatalog = sanitizeIbsherProductCatalog(await getIbsherProductCatalog(options));

  sendToRenderer("ai-product:progress", { step: "analyzing", completed: 0, total: imagePaths.length });
  const analyzed = await analyzeProductImages(imagePaths, options.hints || {}, settings, {}, ibsherCatalog);
  const ai = analyzed.product;
  const keywordsEn = normalizeKeywords(ai.keywordsEn);
  const keywordsAr = normalizeKeywords(ai.keywordsAr);
  const keywordsKu = normalizeKeywords(ai.keywordsKu);
  const brand = cleanCell(options.hints?.brand || ai.brand);
  const collectionName = ibsherCatalog.collections.length
    ? resolveIbsherCollectionName(options.hints?.collectionName, ibsherCatalog)
      || resolveIbsherCollectionName(ai.collectionName, ibsherCatalog)
    : cleanCell(options.hints?.collectionName);
  const row = {
    itemcode: cleanCell(options.hints?.itemcode || ai.itemcode),
    en: removeBrandFromProductName(ai.en, brand),
    ar: removeBrandFromProductName(ai.ar, brand),
    ku: removeBrandFromProductName(ai.ku, brand),
    "description-en": normalizeDescriptionText(ai.descriptionEn),
    "description-ar": normalizeDescriptionText(ai.descriptionAr),
    "description-ku": normalizeDescriptionText(ai.descriptionKu),
    collectionName,
    brand,
    keyword: formatMultilingualKeywords(keywordsEn, keywordsAr, keywordsKu),
    color: normalizeSingleColor(ai.color),
    size: normalizeProductSize(ai.size) || "One size",
  };
  imagePaths.forEach((imagePath, index) => { row[`image-${index + 1}`] = imagePath; });
  sendToRenderer("ai-product:progress", { step: "complete", completed: imagePaths.length, total: imagePaths.length });
  const result = {
    row,
    columns: [...AI_PRODUCT_FIELDS, ...imagePaths.map((_, index) => `image-${index + 1}`)],
    api: analyzed.api,
  };
  return result;
}

async function completeAiProductRow(options = {}) {
  requireAuthenticatedUser();
  const sourceRow = options.row && typeof options.row === "object" ? options.row : {};
  const existingRow = normalizeItemcodeField(sourceRow);
  const imagePaths = Array.isArray(options.imagePaths) ? options.imagePaths : [];
  for (const imagePath of imagePaths) {
    if (!isRemoteImageSource(imagePath) && (!fs.existsSync(imagePath) || !getImageMimeType(imagePath))) {
      throw new Error("One of the Excel row images is unavailable or unsupported.");
    }
  }
  if (!imagePaths.length && !Object.values(existingRow).some((value) => cleanCell(value))) {
    throw new Error("The Excel row has no product data or images for AI completion.");
  }

  const settings = readAiSettings();
  if (!settings.geminiApiKey) throw new Error("Configure the Google AI Studio API key first.");
  const ibsherCatalog = sanitizeIbsherProductCatalog(await getIbsherProductCatalog(options));
  sendToRenderer("ai-product:progress", { step: "analyzing", completed: 0, total: Math.max(1, imagePaths.length) });
  const analyzed = await analyzeProductImages(imagePaths, {}, settings, existingRow, ibsherCatalog);
  const ai = analyzed.product;
  const brand = cleanCell(existingRow.brand || ai.brand);
  const translateExistingProductName = getLocalizedGroupState(existingRow, "name").hasExistingText;
  const normalizeCompletedProductName = (value) => (
    translateExistingProductName ? cleanCell(value) : removeBrandFromProductName(value, brand)
  );
  const collectionName = ibsherCatalog.collections.length
    ? resolveIbsherCollectionName(existingRow.collectionName, ibsherCatalog)
      || resolveIbsherCollectionName(ai.collectionName, ibsherCatalog)
    : cleanCell(existingRow.collectionName);
  const generatedRow = {
    itemcode: cleanCell(ai.itemcode),
    en: normalizeCompletedProductName(ai.en),
    ar: normalizeCompletedProductName(ai.ar),
    ku: normalizeCompletedProductName(ai.ku),
    "description-en": normalizeDescriptionText(ai.descriptionEn),
    "description-ar": normalizeDescriptionText(ai.descriptionAr),
    "description-ku": normalizeDescriptionText(ai.descriptionKu),
    collectionName,
    brand,
    keyword: formatMultilingualKeywords(
      normalizeKeywords(ai.keywordsEn),
      normalizeKeywords(ai.keywordsAr),
      normalizeKeywords(ai.keywordsKu)
    ),
    color: normalizeSingleColor(ai.color),
    size: normalizeProductSize(ai.size) || "One size",
  };
  const completedRow = { ...existingRow };
  for (const field of AI_PRODUCT_FIELDS) {
    if (!cleanCell(completedRow[field])) completedRow[field] = generatedRow[field] || "";
  }
  for (const field of ["description-en", "description-ar", "description-ku"]) {
    completedRow[field] = normalizeDescriptionText(completedRow[field]);
  }
  completedRow.collectionName = collectionName;
  completedRow.color = normalizeSingleColor(completedRow.color);
  completedRow.size = normalizeProductSize(completedRow.size) || "One size";
  imagePaths.forEach((imagePath, index) => {
    completedRow[`image-${index + 1}`] = cleanCell(existingRow[`image-${index + 1}`]) || imagePath;
  });
  const existingImageColumns = Object.keys(completedRow)
    .filter((field) => /^image-\d+$/.test(field) && cleanCell(completedRow[field]));
  if (!existingImageColumns.length) {
    sendToRenderer("ai-product:progress", {
      step: "finding-image",
      completed: 0,
      total: 1,
      itemcode: completedRow.itemcode,
    });
    completedRow["image-1"] = await findExactProductImageForAiRow(completedRow, settings);
  }
  const imageColumns = Object.keys(completedRow)
    .filter((field) => /^image-\d+$/.test(field) && cleanCell(completedRow[field]))
    .sort((left, right) => Number(left.slice(6)) - Number(right.slice(6)));
  sendToRenderer("ai-product:progress", { step: "complete", completed: imageColumns.length, total: Math.max(1, imageColumns.length) });
  const result = {
    row: completedRow,
    columns: [...AI_PRODUCT_FIELDS, ...imageColumns],
    api: analyzed.api,
  };
  return result;
}

async function exportAiProductExcel(options = {}) {
  requireAuthenticatedUser();
  const rows = Array.isArray(options.rows)
    ? options.rows.filter((row) => row && typeof row === "object")
    : options.row && typeof options.row === "object" ? [options.row] : [];
  if (!rows.length) throw new Error("Generate at least one product before exporting.");
  const normalizedRows = rows.map((row) => {
    const normalizedRow = normalizeItemcodeField(row);
    return {
      ...normalizedRow,
      color: normalizeSingleColor(normalizedRow.color),
      size: normalizeProductSize(normalizedRow.size) || "One size",
    };
  });
  const imageColumns = [...new Set(normalizedRows.flatMap((row) => Object.keys(row)))]
    .filter((key) => /^image-\d+$/.test(key))
    .sort((a, b) => Number(a.slice(6)) - Number(b.slice(6)));
  const extraColumns = [...new Set(normalizedRows.flatMap((row) => Object.keys(row)))]
    .filter((key) => !AI_PRODUCT_FIELDS.includes(key) && !/^image-\d+$/.test(key));
  const columns = [...AI_PRODUCT_FIELDS, ...extraColumns, ...imageColumns];
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Download AI products Excel",
    defaultPath: `ai-products-${Date.now()}.xlsx`,
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (result.canceled || !result.filePath) return { canceled: true };
  const xlsx = loadXlsx();
  const worksheet = xlsx.utils.json_to_sheet(normalizedRows, { header: columns });
  worksheet["!cols"] = columns.map((column) => ({ wch: column.startsWith("description") ? 60 : column.startsWith("image-") ? 45 : 24 }));
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Products");
  xlsx.writeFile(workbook, result.filePath);
  return { canceled: false, filePath: result.filePath };
}

function cleanCell(value) {
  return String(value ?? "").trim();
}

function normalizeSingleColor(value) {
  const rawColor = Array.isArray(value) ? value[0] : value;
  return cleanCell(rawColor)
    .replace(/^(?:color|colour)\s*:\s*/i, "")
    .split(/[\n\r,;|/+&]+|\s+(?:and|with|\u0648)\s+/i)[0]
    .trim();
}

function normalizeProductSize(value) {
  const rawSize = Array.isArray(value) ? value[0] : value;
  const size = cleanCell(rawSize)
    .replace(/^size\s*:\s*/i, "")
    .split(/[\n\r,;|/]+|\s+or\s+/i)[0]
    .trim();

  if (/^(?:one[\s-]?size|free[\s-]?size|universal)$/i.test(size)) return "One size";
  if (/^(?:xxs|xs|s|m|l|xl|xxl|xxxl|[2-9]xl)$/i.test(size)) return size.toUpperCase();
  return size;
}

function getExcelRowNumber(row, index) {
  return Number.isInteger(row.__rowNum__) ? row.__rowNum__ + 1 : index + 2;
}

function validateProductRows(rows, sourceRows = rows) {
  const itemCounts = rows.reduce((counts, row) => {
    const itemcode = cleanCell(row.itemcode).toLocaleLowerCase();
    if (itemcode) counts.set(itemcode, (counts.get(itemcode) || 0) + 1);
    return counts;
  }, new Map());
  const conflictsByItemcode = findProductRowConflicts(rows);

  return rows.map((row, index) => {
    const errors = [];
    const warnings = [];
    const itemcode = cleanCell(row.itemcode);
    const itemcodeKey = itemcode.toLocaleLowerCase();
    const price = Number(cleanCell(row.price).replace(/,/g, ""));
    const stockQuantityText = cleanCell(row.stockQuantity);
    const stockQuantity = Number(stockQuantityText.replace(/,/g, ""));
    const isOutOfStock = /^out[\s_-]+of[\s_-]+stock$/i.test(stockQuantityText);
    const imageCount = INTERNAL_PRODUCT_COLUMNS
      .filter((column) => column.startsWith("image-"))
      .filter((column) => cleanCell(row[column])).length;

    for (const column of REQUIRED_PRODUCT_COLUMNS) {
      if (!cleanCell(row[column])) {
        errors.push(`${column} is required`);
      }
    }

    if (row.__priceCurrencyConflict) {
      errors.push("Enter a value in only one of price USD or price IQD");
    } else if (!cleanCell(row.price)) {
      errors.push("price USD or price IQD is required");
    }

    if (itemcode && itemCounts.get(itemcodeKey) > 1) {
      warnings.push("Duplicate itemcode will be uploaded as a color option");
      if (!cleanCell(row.color)) errors.push("color is required for a color option");
      if (!cleanCell(row.size)) errors.push("size is required for a color option");
      if (imageCount === 0) errors.push("At least one image is required for each color option");
    }
    const conflictingColumns = conflictsByItemcode.get(itemcodeKey) || [];
    if (conflictingColumns.length) {
      errors.push(`Rows with the same itemcode must use the same product details: ${conflictingColumns.join(", ")}`);
    }

    if (cleanCell(row.price) && (!Number.isFinite(price) || price <= 0)) {
      errors.push("price must be greater than 0");
    }

    if (stockQuantityText && !isOutOfStock && (!Number.isInteger(stockQuantity) || stockQuantity < 0)) {
      errors.push("stockQuantity must be a whole number of 0 or more, or Out Of Stock");
    }

    if (imageCount === 0 && itemCounts.get(itemcodeKey) <= 1) {
      warnings.push("No images provided");
    }

    return {
      rowId: `${Date.now()}-${index}`,
      rowNumber: getExcelRowNumber(sourceRows[index] || row, index),
      errors,
      warnings,
    };
  });
}

function normalizeProductRows(rows) {
  return rows.map((row) => {
    const sourceRow = normalizeProductPricing(normalizeItemcodeField(row));
    const normalized = {};
    for (const column of INTERNAL_PRODUCT_COLUMNS) {
      normalized[column] = sourceRow[column] ?? "";
    }
    normalized.currency = sourceRow.currency;
    normalized.price = sourceRow.price;
    normalized.__priceCurrencyConflict = sourceRow.__priceCurrencyConflict;
    return normalized;
  });
}

function summarizeProducts(rows, validation) {
  return {
    total: rows.length,
    valid: validation.filter((item) => item.errors.length === 0).length,
    invalid: validation.filter((item) => item.errors.length > 0).length,
    warnings: validation.filter((item) => item.warnings.length > 0).length,
  };
}

async function createExcelTemplate() {
  const xlsx = loadXlsx();
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Save IBSHER Excel template",
    defaultPath: "ibsher-product-template.xlsx",
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  const sampleRow = {
    itemCode: "SKU-1001",
    en: "English product name",
    ar: "Arabic product name",
    ku: "Kurdish product name",
    "description-en": "English description",
    "description-ar": "Arabic description",
    "description-ku": "Kurdish description",
    collectionName: "Collection name",
    brand: "Brand name",
    "price USD": 100,
    "price IQD": "",
    discountPrice: "",
    keyword: "product keyword",
    warrantyDuration: "",
    warrantyUnit: "",
    warrantyDescription: "",
    color: "Black",
    size: "One size",
    stockQuantity: 10,
    "image-1": "https://example.com/black-image.jpg",
  };
  const secondColorRow = {
    ...sampleRow,
    color: "White",
    stockQuantity: 8,
    "image-1": "https://example.com/white-image.jpg",
  };

  const worksheet = xlsx.utils.json_to_sheet([sampleRow, secondColorRow], { header: TEMPLATE_COLUMNS });
  worksheet["!cols"] = TEMPLATE_COLUMNS.map((column) => ({
    wch: Math.max(14, column.length + 4),
  }));

  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  xlsx.writeFile(workbook, result.filePath);

  return { canceled: false, filePath: result.filePath };
}

function parseProductWorkbook(options = {}) {
  const xlsx = loadXlsx();
  const excelFile = String(options.excelFile || "").trim();
  const sheetName = String(options.sheetName || "Sheet1").trim() || "Sheet1";

  if (!excelFile || !fs.existsSync(excelFile)) {
    throw new Error("Excel file was not found.");
  }

  const workbook = xlsx.readFile(excelFile);
  const worksheet = workbook.Sheets[sheetName] || workbook.Sheets[workbook.SheetNames[0]];
  if (!worksheet) {
    throw new Error("No worksheet was found in the selected file.");
  }

  const rawRows = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  const sortedRawRows = sortProductRowsByItemcode(rawRows);
  const rows = normalizeProductRows(sortedRawRows);
  const validation = validateProductRows(rows, sortedRawRows);
  const brands = [...new Set(rows.map((row) => cleanCell(row.brand)).filter(Boolean))].sort();
  const collections = [...new Set(rows.map((row) => cleanCell(row.collectionName)).filter(Boolean))].sort();

  return {
    rows: rows.map((row, index) => ({
      ...row,
      __rowId: validation[index].rowId,
      __rowNumber: validation[index].rowNumber,
      __errors: validation[index].errors,
      __warnings: validation[index].warnings,
    })),
    columns: TEMPLATE_COLUMNS,
    sheetName: worksheet === workbook.Sheets[sheetName] ? sheetName : workbook.SheetNames[0],
    brands,
    collections,
    summary: summarizeProducts(rows, validation),
  };
}

function saveManagedProducts(options = {}) {
  const xlsx = loadXlsx();
  const rows = Array.isArray(options.rows) ? options.rows : [];
  if (!rows.length) throw new Error("There are no product rows left to upload. Restore or import product data first.");
  const sheetName = String(options.sheetName || "Sheet1").trim() || "Sheet1";
  const runDir = String(options.runDir || app.getPath("userData")).trim();
  const outputDir = fs.existsSync(runDir) ? runDir : app.getPath("userData");
  const filePath = path.join(outputDir, `managed-products-${Date.now()}.xlsx`);
  const cleanRows = sortProductRowsByItemcode(rows).map((row) => (
    Object.fromEntries(TEMPLATE_COLUMNS.map((column) => [
      column,
      column === "itemCode" ? row.itemcode ?? row.itemCode ?? "" : row[column] ?? "",
    ]))
  ));

  const worksheet = xlsx.utils.json_to_sheet(cleanRows, { header: TEMPLATE_COLUMNS });
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
  xlsx.writeFile(workbook, filePath);

  return { filePath };
}

function normalizeOptions(options = {}) {
  const excelFile = String(options.excelFile || "").trim();
  const sourceExcelFile = String(options.sourceExcelFile || excelFile).trim();
  const imageFolder = String(options.imageFolder || "").trim();
  const temporaryManagedFile = String(options.temporaryManagedFile || "").trim();
  const sheetName = String(options.sheetName || "Sheet1").trim() || "Sheet1";
  const automationMode = options.automationMode === "browser" ? "browser" : "api";
  const submissionMode = options.submissionMode === "auto" ? "auto" : "manual";
  const email = String(options.email || "").trim();
  const password = String(options.password || "");
  const maxImagesPerProduct = Number.parseInt(options.maxImagesPerProduct || "15", 10) || 15;
  const runDir = String(options.runDir || "").trim() || path.dirname(excelFile);

  return {
    excelFile,
    sourceExcelFile,
    imageFolder,
    temporaryManagedFile,
    sheetName,
    automationMode,
    submissionMode,
    email,
    password,
    maxImagesPerProduct: Math.max(1, maxImagesPerProduct),
    runDir,
  };
}

function validateAutomationOptions(options) {
  if (!options.excelFile) {
    throw new Error("Choose an Excel file before starting.");
  }

  if (!fs.existsSync(options.excelFile)) {
    throw new Error("The selected Excel file could not be found.");
  }

  if (options.imageFolder) {
    if (!fs.existsSync(options.imageFolder) || !fs.statSync(options.imageFolder).isDirectory()) {
      throw new Error("The selected image folder could not be found.");
    }
  }

  if (!options.email || !options.password) {
    throw new Error("IBSHER email and password are required.");
  }

  if (!fs.existsSync(options.runDir)) {
    fs.mkdirSync(options.runDir, { recursive: true });
  }
}

function cleanupRunImageDirectory(downloadDir, runDir) {
  const resolvedRunDir = path.resolve(runDir);
  const resolvedDownloadDir = path.resolve(downloadDir);
  const isExpectedDirectory = path.dirname(resolvedDownloadDir) === resolvedRunDir
    && path.basename(resolvedDownloadDir).toLowerCase() === "temp_images";

  if (isExpectedDirectory && fs.existsSync(resolvedDownloadDir)) {
    fs.rmSync(resolvedDownloadDir, { recursive: true, force: true });
  }
}

function cleanupTemporaryManagedFile(options) {
  if (!options.temporaryManagedFile) return;

  const resolvedManagedFile = path.resolve(options.temporaryManagedFile);
  const isCurrentInput = resolvedManagedFile === path.resolve(options.excelFile);
  const isInRunDirectory = path.dirname(resolvedManagedFile) === path.resolve(options.runDir);
  const isManagedWorkbook = /^managed-products-\d+\.xlsx$/i.test(path.basename(resolvedManagedFile));

  if (isCurrentInput && isInRunDirectory && isManagedWorkbook && fs.existsSync(resolvedManagedFile)) {
    fs.unlinkSync(resolvedManagedFile);
  }
}

function createFallbackAutomationReport(options, details = {}) {
  return {
    generatedAt: new Date().toISOString(),
    excelFile: options.sourceExcelFile,
    sheetName: options.sheetName,
    totalRows: null,
    processed: 0,
    failed: 1,
    imageDownloadFailures: 0,
    incomplete: true,
    process: details,
    failedItems: {
      otherErrors: [{
        itemcode: "(run)",
        name: "Automation run",
        reason: details.message || "The automation ended before it could create the full report.",
      }],
    },
  };
}

async function exportAutomationReport(options = {}) {
  const report = options && typeof options.report === "object" && !Array.isArray(options.report)
    ? options.report
    : null;
  if (!report) {
    throw new Error("No automation report is available to export.");
  }

  const reportDate = new Date(report.generatedAt || Date.now());
  const safeTimestamp = (Number.isNaN(reportDate.getTime()) ? new Date() : reportDate)
    .toISOString()
    .replace(/[:.]/g, "-");
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Export automation report to Excel",
    defaultPath: `automation-report-${safeTimestamp}.xlsx`,
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (result.canceled || !result.filePath) return { canceled: true };

  const xlsx = loadXlsx();
  const exportData = automationReportExport.buildAutomationReportExportData(report);
  const workbook = automationReportExport.createAutomationReportWorkbook(xlsx, report);
  xlsx.writeFile(workbook, result.filePath, { cellStyles: true });

  return {
    canceled: false,
    filePath: result.filePath,
    exportedProducts: exportData.productCount,
    exportedRows: exportData.rows.length,
  };
}

async function startAutomation(rawOptions) {
  await ensureAuthSession();
  requireAuthenticatedUser();

  if (automationProcess) {
    throw new Error("Automation is already running.");
  }

  const vendorAccount = await resolveIbsherVendorAccount(rawOptions);
  const options = normalizeOptions({
    ...rawOptions,
    email: vendorAccount.email,
    password: vendorAccount.password,
  });
  validateAutomationOptions(options);

  const scriptPath = getAutomationScriptPath();
  if (!fs.existsSync(scriptPath)) {
    throw new Error(`Automation script not found: ${scriptPath}`);
  }

  const downloadDir = path.join(options.runDir, "temp_images");
  cleanupRunImageDirectory(downloadDir, options.runDir);
  const inheritedEnvironment = Object.fromEntries(
    Object.entries(process.env).filter(([key]) => !(
      /^REACT_APP_ACCOUNT_\d+_(LABEL|EMAIL|PASSWORD)$/i.test(key)
      || /^IBSHER_(ADMIN|CATEGORY_VENDOR)_(EMAIL|PASSWORD)$/i.test(key)
    )),
  );
  const childEnv = {
    ...inheritedEnvironment,
    ELECTRON_RUN_AS_NODE: "1",
    EXCEL_FILE: options.excelFile,
    SOURCE_EXCEL_FILE: options.sourceExcelFile,
    IMAGE_SOURCE_DIR: options.imageFolder,
    SHEET_NAME: options.sheetName,
    RUN_DIR: options.runDir,
    DOWNLOAD_DIR: downloadDir,
    MAX_IMAGES_PER_PRODUCT: String(options.maxImagesPerProduct),
    AUTOMATION_MODE: options.automationMode,
    SUBMISSION_MODE: options.submissionMode,
    IBSHER_EMAIL: options.email,
    IBSHER_PASSWORD: options.password,
    FORCE_COLOR: "0",
  };

  automationProcess = spawn(process.execPath, [scriptPath], {
    cwd: options.runDir,
    env: childEnv,
    stdio: ["ignore", "pipe", "pipe"],
    windowsHide: false,
  });

  sendToRenderer("automation:status", {
    running: true,
  });

  let automationReport = null;
  let stdoutBuffer = "";
  let pendingStdoutLog = "";
  let stdoutLogTimer = null;

  const flushStdoutLog = () => {
    if (stdoutLogTimer) {
      clearTimeout(stdoutLogTimer);
      stdoutLogTimer = null;
    }
    if (!pendingStdoutLog) return;
    sendToRenderer("automation:log", {
      stream: "stdout",
      text: pendingStdoutLog,
    });
    pendingStdoutLog = "";
  };

  const queueStdoutLog = (text) => {
    if (!text) return;
    pendingStdoutLog += text;
    if (pendingStdoutLog.length >= 32768) {
      flushStdoutLog();
      return;
    }
    if (!stdoutLogTimer) {
      stdoutLogTimer = setTimeout(flushStdoutLog, 80);
    }
  };

  const handleStdoutLine = (line, includeLineBreak = true) => {
    const markerIndex = line.indexOf(AUTOMATION_REPORT_PREFIX);

    if (markerIndex < 0) {
      queueStdoutLog(`${line}${includeLineBreak ? "\n" : ""}`);
      return;
    }

    const visibleText = line.slice(0, markerIndex);
    if (visibleText) {
      queueStdoutLog(`${visibleText}${includeLineBreak ? "\n" : ""}`);
    }

    try {
      const reportJson = line.slice(markerIndex + AUTOMATION_REPORT_PREFIX.length).trim();
      automationReport = JSON.parse(reportJson);
    } catch (error) {
      sendToRenderer("automation:log", {
        stream: "stderr",
        text: `Could not read the automation report: ${error.message}\n`,
      });
    }
  };

  automationProcess.stdout.on("data", (chunk) => {
    stdoutBuffer += chunk.toString();
    const lines = stdoutBuffer.split(/\n/);
    stdoutBuffer = lines.pop() || "";
    lines.forEach((line) => handleStdoutLine(line.replace(/\r$/, "")));
  });

  automationProcess.stderr.on("data", (chunk) => {
    sendToRenderer("automation:log", {
      stream: "stderr",
      text: chunk.toString(),
    });
  });

  automationProcess.on("error", (error) => {
    automationReport = createFallbackAutomationReport(options, { message: error.message });
    cleanupRunImageDirectory(downloadDir, options.runDir);
    cleanupTemporaryManagedFile(options);
    sendToRenderer("automation:log", {
      stream: "stderr",
      text: `${error.message}\n`,
    });
  });

  automationProcess.on("close", (code, signal) => {
    automationProcess = null;
    if (stdoutBuffer) {
      handleStdoutLine(stdoutBuffer.replace(/\r$/, ""), false);
      stdoutBuffer = "";
    }
    flushStdoutLog();
    if (!automationReport) {
      automationReport = createFallbackAutomationReport(options, { exitCode: code, signal });
    }
    cleanupRunImageDirectory(downloadDir, options.runDir);
    cleanupTemporaryManagedFile(options);
    sendToRenderer("automation:status", {
      running: false,
      exitCode: code,
      signal,
      report: automationReport,
    });
  });

  return {
    running: true,
  };
}

function stopAutomation() {
  if (!automationProcess) {
    return { running: false };
  }

  automationProcess.kill();
  automationProcess = null;
  sendToRenderer("automation:status", { running: false, stopped: true });
  return { running: false, stopped: true };
}

let activeIbsherBugScan = null;

function isSafeIbsherCrawlUrl(candidate, origin) {
  try {
    const parsed = new URL(candidate);
    if (parsed.protocol !== "https:" || parsed.origin !== origin) return false;
    if (/\.(?:jpe?g|png|gif|webp|svg|pdf|zip|xlsx?|docx?|mp4|mp3)(?:$|\?)/i.test(parsed.pathname)) return false;
    if (/(?:^|\/)(?:logout|signout|delete|remove|destroy|deactivate)(?:\/|$)/i.test(parsed.pathname)) return false;
    parsed.hash = "";
    return parsed.toString();
  } catch {
    return false;
  }
}

async function loadBugScanPage(scanWindow, pageUrl) {
  let timeout;
  try {
    await Promise.race([
      scanWindow.loadURL(pageUrl, { userAgent: PRODUCT_SEARCH_HEADERS["User-Agent"] }),
      new Promise((resolve, reject) => {
        timeout = setTimeout(() => reject(new Error("The page did not finish loading within 30 seconds.")), 30000);
      }),
    ]);
    await new Promise((resolve) => setTimeout(resolve, 1400));
  } finally {
    clearTimeout(timeout);
  }
}

async function loginBugScanAdmin(scanWindow, credentials) {
  const apiSession = await loginIbsherAdmin(credentials, true);
  const browserSession = scanWindow.webContents.session;
  const apiUrl = getIbsherAdminApiBaseUrl();
  await Promise.all([...apiSession.cookies.entries()].map(([name, value]) => (
    browserSession.cookies.set({ url: apiUrl, name, value, secure: true, sameSite: "lax" })
  )));

  const protectedUrl = "https://marketplace.ibsher.com/add-product";
  await loadBugScanPage(scanWindow, protectedUrl);
  const sessionState = await scanWindow.webContents.executeJavaScript(`(() => {
    const input = document.querySelector('input[type="password"]');
    const visible = Boolean(input && input.getClientRects().length && getComputedStyle(input).visibility !== 'hidden' && getComputedStyle(input).display !== 'none');
    return { visiblePassword: visible, url: location.href };
  })()`, true);
  if (!sessionState.visiblePassword) return;

  const loginState = await scanWindow.webContents.executeJavaScript(`(() => ({
    hasEmail: Boolean(document.querySelector('input[type="email"]')),
    hasPassword: Boolean(document.querySelector('input[type="password"]')),
  }))()`, true);
  if (!loginState.hasEmail || !loginState.hasPassword) return;

  const payload = JSON.stringify({ email: credentials.email, password: credentials.password });
  const submitted = await scanWindow.webContents.executeJavaScript(`(() => {
    const credentials = ${payload};
    const email = document.querySelector('input[type="email"]');
    const password = document.querySelector('input[type="password"]');
    if (!email || !password) return false;
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
    setter.call(email, credentials.email);
    setter.call(password, credentials.password);
    for (const input of [email, password]) {
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
    }
    return true;
  })()`, true);
  if (!submitted) throw new Error("The Admin iBsher login form could not be filled.");

  await new Promise((resolve) => setTimeout(resolve, 350));
  const clicked = await scanWindow.webContents.executeJavaScript(`(() => {
    const password = document.querySelector('input[type="password"]');
    const form = password?.closest('form');
    const submit = form?.querySelector('button[type="submit"], input[type="submit"]') || document.querySelector('button[type="submit"]');
    if (!submit) return false;
    if (form?.requestSubmit) form.requestSubmit(submit);
    else submit.click();
    return true;
  })()`, true);
  if (!clicked) throw new Error("The Admin iBsher login form has no submit button.");

  await new Promise((resolve) => setTimeout(resolve, 2500));
  await loadBugScanPage(scanWindow, protectedUrl);
  const finalState = await scanWindow.webContents.executeJavaScript(`(() => {
    const input = document.querySelector('input[type="password"]');
    const visible = Boolean(input && input.getClientRects().length && getComputedStyle(input).visibility !== 'hidden' && getComputedStyle(input).display !== 'none');
    const error = Array.from(document.querySelectorAll('[role="alert"], .error, .ant-form-item-explain-error'))
      .map((element) => element.textContent.trim()).filter(Boolean).join(' · ').slice(0, 500);
    return { visiblePassword: visible, error, url: location.href };
  })()`, true);
  if (finalState.visiblePassword) {
    throw new Error(finalState.error
      ? `Admin iBsher login failed: ${finalState.error}`
      : "Admin iBsher returned to the login page after authentication. Check that this vendor account can access the marketplace.");
  }
}

async function captureBugScanSnapshot(scanWindow) {
  return scanWindow.webContents.executeJavaScript(`(() => {
    const selector = (element) => {
      if (!element) return '';
      if (element.id) return '#' + CSS.escape(element.id);
      const name = element.tagName ? element.tagName.toLowerCase() : 'element';
      const className = typeof element.className === 'string'
        ? element.className.trim().split(/\\s+/).filter(Boolean).slice(0, 2).map((value) => '.' + CSS.escape(value)).join('') : '';
      return (name + className).slice(0, 180);
    };
    const accessibleName = (element) => (element.getAttribute('aria-label') || element.getAttribute('title') || element.textContent || element.value || '').trim();
    const duplicateIds = Array.from(document.querySelectorAll('[id]')).map((element) => element.id).filter((id, index, ids) => id && ids.indexOf(id) !== index).filter((id, index, ids) => ids.indexOf(id) === index).slice(0, 40);
    const headings = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,h6')).map((element) => ({ level: Number(element.tagName.slice(1)), text: element.textContent.trim().slice(0, 90) }));
    const headingSkips = headings.filter((heading, index) => index > 0 && heading.level > headings[index - 1].level + 1).map((heading, index) => 'H' + heading.level + ': ' + (heading.text || 'untitled heading')).slice(0, 30);
    const navigation = performance.getEntriesByType('navigation')[0];
    const resources = performance.getEntriesByType('resource');
    const origin = location.origin;
    const unsafePath = /(?:^|\\/)(?:logout|signout|delete|remove|destroy|deactivate)(?:\\/|$)/i;
    const internalLinks = Array.from(document.querySelectorAll('a[href]')).map((link) => {
      try { const parsed = new URL(link.href, location.href); parsed.hash = ''; return parsed.origin === origin && !unsafePath.test(parsed.pathname) ? parsed.href : ''; } catch { return ''; }
    }).filter(Boolean).filter((value, index, values) => values.indexOf(value) === index).slice(0, 120);
    return {
      title: document.title.trim(),
      description: document.querySelector('meta[name="description"]')?.content?.trim() || '',
      canonical: document.querySelector('link[rel="canonical"]')?.href || '',
      viewport: document.querySelector('meta[name="viewport"]')?.content || '',
      lang: document.documentElement.lang || '',
      bodyTextLength: (document.body?.innerText || '').trim().length,
      h1Count: document.querySelectorAll('h1').length,
      hasMainLandmark: Boolean(document.querySelector('main, [role="main"]')),
      horizontalOverflow: document.documentElement.scrollWidth > document.documentElement.clientWidth + 4,
      scrollWidth: document.documentElement.scrollWidth,
      clientWidth: document.documentElement.clientWidth,
      navigationMs: navigation?.loadEventEnd || navigation?.duration || 0,
      transferBytes: resources.reduce((total, item) => total + Number(item.transferSize || 0), Number(navigation?.transferSize || 0)),
      brokenImages: Array.from(document.images).filter((image) => image.complete && image.naturalWidth === 0).map((image) => image.currentSrc || image.src || selector(image)).filter(Boolean).slice(0, 40),
      missingAltImages: Array.from(document.images).filter((image) => !image.hasAttribute('alt') && image.getAttribute('role') !== 'presentation').map((image) => image.currentSrc || image.src || selector(image)).filter(Boolean).slice(0, 40),
      mixedContent: Array.from(document.querySelectorAll('[src], [href]')).map((element) => element.src || element.href || '').filter((value) => /^http:\\/\\//i.test(value)).slice(0, 40),
      unlabeledControls: Array.from(document.querySelectorAll('button, input, select, textarea, [role="button"]')).filter((element) => {
        if (element.disabled || element.type === 'hidden' || accessibleName(element)) return false;
        const id = element.id;
        return !(id && document.querySelector('label[for="' + CSS.escape(id) + '"]'));
      }).map(selector).filter(Boolean).slice(0, 40),
      invalidLinks: Array.from(document.querySelectorAll('a')).filter((link) => !link.getAttribute('href') || /^(?:#|javascript:)/i.test(link.getAttribute('href'))).map(selector).filter(Boolean).slice(0, 40),
      insecureForms: Array.from(document.forms).map((form) => form.action || '').filter((action) => action && !/^https:\\/\\//i.test(action)).slice(0, 30),
      duplicateIds,
      headingSkips,
      internalLinks,
    };
  })()`, true);
}

async function scanIbsherTarget(targetInput, credentials, maxPages) {
  const target = ibsherBugFinder.normalizeScanTarget(targetInput);
  const origin = new URL(target.url).origin;
  const startedAt = Date.now();
  const networkEvents = [];
  const consoleEvents = [];
  const reports = [];
  const visited = new Set();
  const queued = new Set();
  const queue = [];
  const partition = `ibsher-bug-scan-${startedAt}-${crypto.randomBytes(5).toString("hex")}`;
  let scanWindow;
  let currentPageUrl = target.url;
  let mainStatus = null;

  const enqueue = (candidate) => {
    const safeUrl = isSafeIbsherCrawlUrl(candidate, origin);
    if (!safeUrl || visited.has(safeUrl) || queued.has(safeUrl)) return;
    queued.add(safeUrl);
    queue.push(safeUrl);
  };

  try {
    scanWindow = new BrowserWindow({
      show: false,
      width: 1366,
      height: 900,
      webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true, javascript: true, partition },
    });
    activeIbsherBugScan.window = scanWindow;
    scanWindow.setMenu(null);
    scanWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
    const scanSession = scanWindow.webContents.session;
    scanSession.setPermissionRequestHandler((webContents, permission, callback) => callback(false));
    const requestFilter = { urls: ["http://*/*", "https://*/*"] };
    scanSession.webRequest.onCompleted(requestFilter, (details) => {
      if (details.resourceType === "mainFrame") mainStatus = details.statusCode;
      if (details.statusCode >= 400) networkEvents.push({ pageUrl: currentPageUrl, status: details.statusCode, resourceType: details.resourceType, url: details.url });
    });
    scanSession.webRequest.onErrorOccurred(requestFilter, (details) => {
      if (details.error !== "net::ERR_ABORTED") networkEvents.push({ pageUrl: currentPageUrl, status: 0, resourceType: details.resourceType, url: details.url, error: details.error });
    });
    scanWindow.webContents.on("console-message", (event, level, message, line, sourceId) => {
      if (level >= 2 && String(message || "").trim()) consoleEvents.push({ pageUrl: currentPageUrl, level, message, line, sourceId });
    });

    currentPageUrl = target.url;
    await loadBugScanPage(scanWindow, target.url);
    if (target.authenticated) {
      if (!credentials) throw new Error("Choose a saved vendor account for the Admin iBsher scan.");
      sendToRenderer("ibsher-bug-finder:progress", { phase: "login", target: target.name, message: "Signing in to Admin iBsher..." });
      await loginBugScanAdmin(scanWindow, credentials);
    }
    enqueue(scanWindow.webContents.getURL() || target.url);

    while (queue.length && reports.length < maxPages) {
      if (activeIbsherBugScan?.canceled) throw new Error("Bug scan canceled.");
      const pageUrl = queue.shift();
      queued.delete(pageUrl);
      if (visited.has(pageUrl)) continue;
      visited.add(pageUrl);
      currentPageUrl = pageUrl;
      mainStatus = null;
      const pageStartedAt = Date.now();
      const networkStart = networkEvents.length;
      const consoleStart = consoleEvents.length;
      const findings = [];
      sendToRenderer("ibsher-bug-finder:progress", { phase: "scan", target: target.name, current: reports.length + 1, maxPages, url: pageUrl, message: `Scanning ${target.name} page ${reports.length + 1} of ${maxPages}...` });

      let snapshot = null;
      try {
        if (scanWindow.webContents.getURL() !== pageUrl) await loadBugScanPage(scanWindow, pageUrl);
        snapshot = await captureBugScanSnapshot(scanWindow);
        findings.push(...ibsherBugFinder.buildPageFindings(snapshot));
        for (const link of snapshot.internalLinks || []) enqueue(link);
      } catch (error) {
        findings.push(ibsherBugFinder.createFinding("critical", "Availability", "The page could not be loaded or inspected.", error.message));
      }

      if (mainStatus >= 400) findings.push(ibsherBugFinder.createFinding("critical", "HTTP", `The main page returned HTTP ${mainStatus}.`, pageUrl));
      for (const failure of networkEvents.slice(networkStart, networkStart + 100)) {
        const severity = failure.resourceType === "mainFrame" ? "critical" : "error";
        const result = failure.status ? `HTTP ${failure.status}` : failure.error || "Request failed";
        findings.push(ibsherBugFinder.createFinding(severity, "Network", `${failure.resourceType || "Resource"} request failed: ${result}.`, failure.url));
      }
      for (const item of consoleEvents.slice(consoleStart, consoleStart + 80)) {
        findings.push(ibsherBugFinder.createFinding(item.level >= 3 ? "error" : "warning", "Console", item.message, `${item.sourceId || "page"}${item.line ? `:${item.line}` : ""}`));
      }
      const unique = [...new Map(findings.map((finding) => [`${finding.severity}|${finding.category}|${finding.message}|${finding.details}`, finding])).values()].slice(0, 300);
      reports.push({
        name: target.name,
        authenticated: target.authenticated,
        url: pageUrl,
        finalUrl: scanWindow.webContents.getURL() || pageUrl,
        title: snapshot?.title || "",
        scannedAt: new Date().toISOString(),
        durationMs: Date.now() - pageStartedAt,
        status: mainStatus,
        findings: unique,
        summary: ibsherBugFinder.summarizeFindings(unique),
      });
    }
    return reports;
  } finally {
    if (scanWindow && !scanWindow.isDestroyed()) scanWindow.destroy();
    if (activeIbsherBugScan) activeIbsherBugScan.window = null;
  }
}

async function runIbsherBugScan(options = {}) {
  requireAccountSectionAccess("bug-finder");
  if (activeIbsherBugScan) throw new Error("An iBsher bug scan is already running.");
  const targets = ibsherBugFinder.normalizeScanTargets(options.targets);
  const maxPages = Math.max(1, Math.min(25, Number.parseInt(options.maxPages, 10) || 8));
  activeIbsherBugScan = { canceled: false, window: null };
  try {
    const credentials = targets.some((target) => target?.authenticated)
      ? await resolveIbsherVendorAccount({ vendorAccountValue: options.vendorAccountValue })
      : null;
    const reports = [];
    for (const target of targets) reports.push(...await scanIbsherTarget(target, credentials, maxPages));
    const summary = ibsherBugFinder.summarizeFindings(reports.flatMap((report) => report.findings));
    sendToRenderer("ibsher-bug-finder:progress", { phase: "complete", message: "Detailed bug scan complete.", pages: reports.length, summary });
    return { scannedAt: new Date().toISOString(), reports, summary };
  } finally {
    activeIbsherBugScan = null;
  }
}

function cancelIbsherBugScan() {
  if (!activeIbsherBugScan) return { cancelRequested: false };
  activeIbsherBugScan.canceled = true;
  activeIbsherBugScan.window?.webContents?.stop();
  return { cancelRequested: true };
}

function registerProtectedIpcHandler(channel, handler, requiredSections = null) {
  ipcMain.handle(channel, async (event, ...args) => {
    await validateAuthSession();
    if (requiredSections) requireAccountSectionAccess(requiredSections);
    return handler(event, ...args);
  });
}

app.whenReady().then(() => {
  removeLegacyLocalVendorCredentialStores();
  registerProtectedIpcHandler("dialog:select-excel-file", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Choose product spreadsheet",
      properties: ["openFile"],
      filters: [
        { name: "Spreadsheets", extensions: ["xlsx", "xls", "csv"] },
        { name: "All Files", extensions: ["*"] },
      ],
    });

    return result;
  }, ["uploader", "products"]);

  registerProtectedIpcHandler("dialog:select-product-image-folder", async () => {
    return dialog.showOpenDialog(mainWindow, {
      title: "Choose product image folder",
      properties: ["openDirectory"],
    });
  }, "uploader");

  registerProtectedIpcHandler("products:create-template", () => createExcelTemplate(), ["uploader", "products"]);
  registerProtectedIpcHandler("products:parse-workbook", (event, options) => parseProductWorkbook(options), ["uploader", "products"]);
  registerProtectedIpcHandler("products:save-managed", (event, options) => saveManagedProducts(options), ["uploader", "products"]);
  registerProtectedIpcHandler("ibsher-export:get-filters", (event, options) => getIbsherExportFilters(options));
  registerProtectedIpcHandler("ibsher-exchange:get-status", async () => {
    const service = getExchangeRateService();
    await service.access.refresh();
    return service.status();
  }, "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:setup-pin", async (event, options) => {
    const service = getExchangeRateService();
    await service.access.setup(options?.pin);
    return service.status();
  }, "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:unlock", async (event, options) => {
    const service = getExchangeRateService();
    await service.access.unlock(options?.pin);
    service.prepared = null;
    if (service.state.enabled) return service.check({ reviewCurrentProducts: true });
    return service.status();
  }, "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:lock", async () => {
    const service = getExchangeRateService();
    service.prepared = null;
    await service.access.lock(() => service.whenIdle());
    service.emit();
    return service.status();
  }, "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:check", () => getExchangeRateService().check({ reviewCurrentProducts: true }), "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:preview", (event, options) => getExchangeRateService().preview(options?.vendorValues), "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:enroll", (event, options) => getExchangeRateService().enroll(options?.token, options?.discounts), "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:set-enabled", (event, options) => getExchangeRateService().setEnabled(options?.enabled === true), "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:set-schedule", (event, options) => getExchangeRateService().setSchedule(options), "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:get-discount-products", () => getExchangeRateService().discountProducts(), "exchange-rates");
  registerProtectedIpcHandler("ibsher-exchange:set-product-discounts", (event, options) => getExchangeRateService().setProductDiscounts(options?.discounts), "exchange-rates");
  registerProtectedIpcHandler("ibsher-export:export", (event, options) => exportIbsherProducts(options));
  registerProtectedIpcHandler("ibsher-export:select-duplicate-file", () => selectIbsherDuplicateExcelFile());
  registerProtectedIpcHandler("ibsher-export:clean-duplicates", (event, options) => cleanIbsherDuplicateItemCodes(options));
  registerProtectedIpcHandler("ibsher-itemcode-update:select-file", () => selectIbsherItemcodeUpdateFile(), "ibsher-export");
  registerProtectedIpcHandler("ibsher-itemcode-update:run", (event, options) => updateIbsherItemcodes(options), "ibsher-export");
  registerProtectedIpcHandler("ibsher-bulk-update:select-file", () => selectIbsherBulkUpdateFile(), "ibsher-export");
  registerProtectedIpcHandler("ibsher-bulk-update:run", (event, options) => bulkUpdateIbsherProducts(options), "ibsher-export");
  registerProtectedIpcHandler("ibsher-bulk-update:export-issues", (event, options) => exportIbsherBulkUpdateIssues(options), "ibsher-export");
  registerProtectedIpcHandler("ibsher-audit:list-products", (event, options) => getIbsherAdminProducts(options));
  registerProtectedIpcHandler("ibsher-audit:list-collections", (event, options) => getIbsherAdminCollections(options));
  registerProtectedIpcHandler("ibsher-audit:get-status", () => getIbsherAuditStatus());
  registerProtectedIpcHandler("ibsher-audit:start", (event, options) => startIbsherProductAudit(options));
  registerProtectedIpcHandler("ibsher-audit:cancel", () => cancelIbsherProductAudit());
  registerProtectedIpcHandler("ibsher-audit:remove-solid-backgrounds", (event, options) => removeIbsherSolidBackgrounds(options));
  registerProtectedIpcHandler("ibsher-audit:auto-fix", (event, options) => autoFixIbsherAdminProduct(options));
  registerProtectedIpcHandler("ibsher-audit:update-product", (event, options) => (
    updateIbsherAdminProduct(options?.product || options, options)
  ));
  registerProtectedIpcHandler("ibsher-bug-finder:scan", (event, options) => runIbsherBugScan(options), "bug-finder");
  registerProtectedIpcHandler("ibsher-bug-finder:cancel", () => cancelIbsherBugScan(), "bug-finder");

  registerProtectedIpcHandler("ai-product:get-settings", () => publicAiSettings());
  registerProtectedIpcHandler("ai-product:save-settings", (event, settings) => saveAiSettings(settings));
  registerProtectedIpcHandler("ai-product:get-api-status", () => getAiProductApiStatus(), "ai-generator");
  registerProtectedIpcHandler("ai-product:select-images", () => selectAiProductImages(), "ai-generator");
  registerProtectedIpcHandler("ai-product:select-image-folder", () => selectAiProductImageFolder(), "ai-generator");
  registerProtectedIpcHandler("ai-product:load-image-previews", (event, options) => loadAiProductImagePreviews(options), "ai-generator");
  registerProtectedIpcHandler("ai-product:load-image-preview", (event, options) => loadAiProductImagePreview(options), "ai-generator");
  registerProtectedIpcHandler("ai-product:paste-image", () => pasteAiProductImage(), "ai-generator");
  registerProtectedIpcHandler("ai-product:import-excel", () => importAiProductsFromExcel(), "ai-generator");
  registerProtectedIpcHandler("ai-product:generate", (event, options) => generateAiProduct(options), "ai-generator");
  registerProtectedIpcHandler("ai-product:complete-row", (event, options) => completeAiProductRow(options), "ai-generator");
  registerProtectedIpcHandler("ai-product:export-excel", (event, options) => exportAiProductExcel(options), "ai-generator");
  registerProtectedIpcHandler("postimages:get-settings", () => publicPostimagesSettings(), "postimages-uploader");
  registerProtectedIpcHandler("postimages:save-settings", (event, options) => savePostimagesSettings(options), "postimages-uploader");
  registerProtectedIpcHandler("postimages:select-images", () => selectPostimagesImages(), "postimages-uploader");
  registerProtectedIpcHandler("postimages:select-excel", () => selectPostimagesExcel(), "postimages-uploader");
  registerProtectedIpcHandler("postimages:upload-images", (event, options) => uploadPostimagesImages(options), "postimages-uploader");
  registerProtectedIpcHandler("postimages:upload-excel", (event, options) => uploadPostimagesExcel(options), "postimages-uploader");
  registerProtectedIpcHandler("product-image-search:import-excel", () => importProductImageSearchExcel(), "ai-generator");
  registerProtectedIpcHandler("product-image-search:find", (event, options) => findOfficialProductImages(options), "ai-generator");
  registerProtectedIpcHandler("product-image-search:export-excel", (event, options) => exportProductImageExcel(options), "ai-generator");
  registerProtectedIpcHandler("ibsher-page-images:import-excel", () => importIbsherPageImageExcel(), "page-image-extractor");
  registerProtectedIpcHandler("ibsher-page-images:extract", (event, options) => extractIbsherPageImageLinks(options), "page-image-extractor");
  registerProtectedIpcHandler("ibsher-page-images:export-excel", (event, options) => exportIbsherPageImageExcel(options), "page-image-extractor");

  registerProtectedIpcHandler("ibsher-brands:select-file", () => selectIbsherBrandWorkbook(), "brand-import");
  registerProtectedIpcHandler("ibsher-brands:parse-workbook", (event, options) => parseIbsherBrandWorkbook(options), "brand-import");
  registerProtectedIpcHandler("ibsher-brands:import", (event, options) => runIbsherBrandImport(options), "brand-import");
  registerProtectedIpcHandler("ibsher-brands:cancel", () => {
    if (activeBrandImport) activeBrandImport.canceled = true;
    return { cancelRequested: Boolean(activeBrandImport) };
  }, "brand-import");

  registerProtectedIpcHandler("ibsher-publishing:list", (event, options) => getIbsherPublishingProducts(options), "product-publishing");
  registerProtectedIpcHandler("ibsher-publishing:filters", (event, options) => getIbsherPublishingFilters(options), "product-publishing");
  registerProtectedIpcHandler("ibsher-publishing:update-status", (event, options) => changeIbsherProductStatuses(options), "product-publishing");
  registerProtectedIpcHandler("ibsher-publishing:publish-all", (event, options) => changeIbsherProductStatuses(options, true), "product-publishing");
  registerProtectedIpcHandler("ibsher-publishing:cancel", () => {
    if (activeProductPublishing) activeProductPublishing.canceled = true;
    return { cancelRequested: Boolean(activeProductPublishing) };
  }, "product-publishing");
  registerProtectedIpcHandler("ibsher-integration-review:list", (event, options) => getIbsherPendingVendorImports(options), "integration-review");
  registerProtectedIpcHandler("ibsher-integration-review:export", (event, options) => exportIbsherIntegrationReview(options), "integration-review");
  registerProtectedIpcHandler("ibsher-integration-review:run", (event, options) => processIbsherPendingVendorImports(options), "integration-review");
  registerProtectedIpcHandler("ibsher-integration-review:cancel", () => {
    if (activeIntegrationReview) activeIntegrationReview.canceled = true;
    return { cancelRequested: Boolean(activeIntegrationReview) };
  }, "integration-review");
  registerProtectedIpcHandler("ibsher-product-delete:select-excel", () => selectIbsherProductDeletionExcel(), "integration-review");
  registerProtectedIpcHandler("ibsher-product-delete:preview", (event, options) => previewIbsherVendorProductDeletion(options), "integration-review");
  registerProtectedIpcHandler("ibsher-product-delete:run", (event, options) => deleteIbsherVendorProductsFromExcel(options), "integration-review");
  registerProtectedIpcHandler("ibsher-product-delete:cancel", () => {
    if (activeVendorProductDeletion) activeVendorProductDeletion.canceled = true;
    return { cancelRequested: Boolean(activeVendorProductDeletion) };
  }, "integration-review");

  registerProtectedIpcHandler("automation:start", (event, options) => startAutomation(options), "uploader");
  registerProtectedIpcHandler("automation:stop", () => stopAutomation(), "uploader");
  registerProtectedIpcHandler("automation:export-report", (event, options) => exportAutomationReport(options), "uploader");

  ipcMain.handle("auth:get-session", () => getAuthSession());
  ipcMain.handle("auth:login", (event, credentials) => loginUser(credentials));
  ipcMain.handle("auth:logout", () => logoutUser());

  registerProtectedIpcHandler("accounts:list", () => listAccounts());
  registerProtectedIpcHandler("accounts:select-profile-image", () => selectProfileImage());
  registerProtectedIpcHandler("accounts:save", (event, account) => saveAccount(account));
  registerProtectedIpcHandler("accounts:delete", (event, accountValue) => deleteAccount(accountValue));
  registerProtectedIpcHandler("accounts:check", (event, accountValue) => checkAccount(accountValue));
  registerProtectedIpcHandler("vendor-accounts:list", () => listVendorAccounts());
  registerProtectedIpcHandler("vendor-accounts:save", (event, account) => saveVendorAccount(account));
  registerProtectedIpcHandler("vendor-accounts:delete", (event, accountValue) => deleteVendorAccount(accountValue));
  registerProtectedIpcHandler("vendor-accounts:check", (event, accountValue) => checkVendorAccount(accountValue));
  ipcMain.handle("updater:get-status", () => ({
    currentVersion: APP_VERSION,
    checking: isCheckingForUpdate,
    downloading: isDownloadingUpdate,
    availableUpdate: availableUpdateInfo,
    downloadedUpdate: downloadedUpdateInfo,
  }));

  ipcMain.handle("updater:check", () => checkForUpdates("manual"));
  ipcMain.handle("updater:download", () => downloadAvailableUpdate());
  ipcMain.handle("updater:quit-and-install", () => {
    emitUpdaterStatus("quit-and-install");
    autoUpdater.quitAndInstall(false, true);
  });

  registerProtectedIpcHandler("path:open", async (event, targetPath) => {
    const cleanPath = String(targetPath || "").trim();
    if (!cleanPath) {
      return { opened: false };
    }

    if (fs.existsSync(cleanPath)) {
      const stats = fs.statSync(cleanPath);
      if (stats.isFile()) {
        shell.showItemInFolder(cleanPath);
        return { opened: true };
      }

      const error = await shell.openPath(cleanPath);
      return { opened: !error, error };
    }

    const parentPath = path.dirname(cleanPath);
    if (parentPath && fs.existsSync(parentPath)) {
      const error = await shell.openPath(parentPath);
      return { opened: !error, error };
    }

    const error = await shell.openPath(cleanPath);
    return { opened: !error, error };
  });

  ipcMain.on("open-external-link", (event, externalUrl) => {
    shell.openExternal(externalUrl);
  });

  // Register every renderer API before loading the React application.
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("before-quit", () => {
  exchangeRateService?.access?.lock(() => exchangeRateService.whenIdle()).catch(() => {});
  if (exchangeRateTimer) clearInterval(exchangeRateTimer);
  if (automationProcess) {
    automationProcess.kill();
    automationProcess = null;
  }
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
