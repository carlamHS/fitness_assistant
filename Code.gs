/**
 * Fitness Assistant
 * Google Apps Script backend for workout and protein tracking.
 */

const SHEETS = Object.freeze({
  WORKOUT_LOG: "WorkoutLog",
  FOOD_CATALOG: "FoodCatalog",
  PROTEIN_LOG: "ProteinLog",
  PROFILES: "Profiles",
});

const WORKOUT_TYPES = Object.freeze({
  STRENGTH: "strength",
  CARDIO: "cardio",
});

const DEFAULT_PROFILE_NAME = "Default";

const HEADERS = Object.freeze({
  WORKOUT_LOG: [
    "CreatedAt",
    "Date",
    "ProfileName",
    "Training",
    "WorkoutType",
    "WeightKg",
    "Reps",
    "Sets",
    "DurationMin",
    "Intensity",
    "SpeedKph",
    "Notes",
  ],
  FOOD_CATALOG: ["CreatedAt", "FoodName", "ProteinPer100g"],
  PROTEIN_LOG: [
    "CreatedAt",
    "Date",
    "ProfileName",
    "FoodName",
    "IntakeGrams",
    "ProteinPer100g",
    "ProteinGrams",
    "Notes",
  ],
  PROFILES: ["CreatedAt", "ProfileName"],
});

function doGet() {
  setupSheets();
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Fitness Assistant")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(spreadsheet, SHEETS.WORKOUT_LOG, HEADERS.WORKOUT_LOG);
  ensureSheet(spreadsheet, SHEETS.FOOD_CATALOG, HEADERS.FOOD_CATALOG);
  ensureSheet(spreadsheet, SHEETS.PROTEIN_LOG, HEADERS.PROTEIN_LOG);
  ensureSheet(spreadsheet, SHEETS.PROFILES, HEADERS.PROFILES);
  ensureDefaultProfile_();
}

function initApp(request) {
  setupSheets();
  const profiles = listProfiles_();
  const activeProfile = resolveRequestedProfile_(request && request.profileName, profiles);
  const trainingName = normalizeTrainingNameFilter_(request && request.trainingName);
  const trendDays = normalizeTrendDays_(request && request.trendDays);
  const workoutNames = listWorkoutNames_(activeProfile);
  const today = normalizeDateString();
  return {
    today: today,
    timezone: Session.getScriptTimeZone(),
    profiles: profiles,
    activeProfile: activeProfile,
    foods: listFoods_(),
    workoutNames: workoutNames,
    recentWorkoutNames: listRecentWorkoutNames_(activeProfile, 8),
    workouts: listWorkoutsByDate_(today, activeProfile),
    proteins: listProteinEntriesByDate_(today, activeProfile),
    dailyTotalProtein: sumProteinByDate_(today, activeProfile),
    weeklyTrend: getWeeklyTrend({
      date: today,
      profileName: activeProfile,
      trainingName: trainingName,
      trendDays: trendDays,
    }),
  };
}

function addWorkout(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    setupSheets();
    const profileName = normalizeProfileName_(payload && payload.profileName);
    ensureProfileExistsByName_(profileName);
    payload.profileName = profileName;
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.WORKOUT_LOG);
    const row = parseWorkoutPayload_(payload);
    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();
    return {
      ok: true,
      workout: parseWorkoutRow_(row, rowNumber),
    };
  } finally {
    lock.releaseLock();
  }
}

function updateWorkout(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    setupSheets();
    const profileName = normalizeProfileName_(payload && payload.profileName);
    ensureProfileExistsByName_(profileName);
    payload.profileName = profileName;
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.WORKOUT_LOG);
    const rowNumber = parseRowNumber_(payload && payload.rowNumber, sheet);
    const row = parseWorkoutPayload_(payload);
    const existingCreatedAt = sheet.getRange(rowNumber, 1).getValue() || new Date();
    row[0] = existingCreatedAt;
    sheet.getRange(rowNumber, 1, 1, HEADERS.WORKOUT_LOG.length).setValues([row]);

    return {
      ok: true,
      workout: parseWorkoutRow_(row, rowNumber),
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteWorkout(rowNumberInput) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    setupSheets();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.WORKOUT_LOG);
    const rowNumber = parseRowNumber_(rowNumberInput, sheet);
    sheet.deleteRow(rowNumber);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function addOrUpdateFood(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    setupSheets();
    const clean = parseFoodPayload_(payload);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.FOOD_CATALOG);
    const updated = upsertFood_(sheet, clean.foodName, clean.foodKey, clean.proteinPer100g);

    return {
      ok: true,
      updated: updated,
      foodName: clean.foodName,
      proteinPer100g: clean.proteinPer100g,
    };
  } finally {
    lock.releaseLock();
  }
}

function addProfile(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    setupSheets();
    const profileName = normalizeProfileName_(payload && payload.profileName);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.PROFILES);
    const existing = findProfileName_(sheet, profileName);
    const created = !existing;

    if (created) {
      sheet.appendRow([new Date(), profileName]);
    }

    const profiles = listProfiles_();
    const activeProfile = resolveRequestedProfile_(profileName, profiles);
    return {
      ok: true,
      created: created,
      profiles: profiles,
      activeProfile: activeProfile,
    };
  } finally {
    lock.releaseLock();
  }
}

function getWorkoutNames(request) {
  setupSheets();
  const profiles = listProfiles_();
  const profileName = resolveRequestedProfile_(request && request.profileName, profiles);
  return listWorkoutNames_(profileName);
}

function getRecentWorkoutNames(request) {
  setupSheets();
  const profiles = listProfiles_();
  const profileName = resolveRequestedProfile_(request && request.profileName, profiles);
  const limitInput = request && request.limit;
  const limit = Number.isFinite(Number(limitInput)) ? Number(limitInput) : 8;
  return listRecentWorkoutNames_(profileName, limit);
}

function addProteinEntry(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    setupSheets();
    const profileName = normalizeProfileName_(payload && payload.profileName);
    ensureProfileExistsByName_(profileName);
    payload.profileName = profileName;

    const clean = parseProteinPayload_(payload);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const foodSheet = spreadsheet.getSheetByName(SHEETS.FOOD_CATALOG);
    const proteinSheet = spreadsheet.getSheetByName(SHEETS.PROTEIN_LOG);
    const resolved = resolveProteinPer100g_(clean, foodSheet);
    const usedProteinPer100g = resolved.proteinPer100g;
    const addedFood = resolved.addedFood;

    const proteinGrams = round2_((clean.intakeGrams * usedProteinPer100g) / 100);
    const row = [
      new Date(),
      clean.date,
      clean.profileName,
      clean.foodName,
      clean.intakeGrams,
      usedProteinPer100g,
      proteinGrams,
      clean.notes,
    ];

    proteinSheet.appendRow(row);
    const rowNumber = proteinSheet.getLastRow();

    return {
      ok: true,
      proteinEntry: parseProteinRow_(row, rowNumber),
      addedFood: addedFood,
    };
  } finally {
    lock.releaseLock();
  }
}

function updateProteinEntry(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    setupSheets();
    const profileName = normalizeProfileName_(payload && payload.profileName);
    ensureProfileExistsByName_(profileName);
    payload.profileName = profileName;
    const clean = parseProteinPayload_(payload);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const foodSheet = spreadsheet.getSheetByName(SHEETS.FOOD_CATALOG);
    const proteinSheet = spreadsheet.getSheetByName(SHEETS.PROTEIN_LOG);
    const rowNumber = parseRowNumber_(payload && payload.rowNumber, proteinSheet);

    const resolved = resolveProteinPer100g_(clean, foodSheet);
    const usedProteinPer100g = resolved.proteinPer100g;
    const proteinGrams = round2_((clean.intakeGrams * usedProteinPer100g) / 100);

    const existingCreatedAt = proteinSheet.getRange(rowNumber, 1).getValue() || new Date();
    const row = [
      existingCreatedAt,
      clean.date,
      clean.profileName,
      clean.foodName,
      clean.intakeGrams,
      usedProteinPer100g,
      proteinGrams,
      clean.notes,
    ];

    proteinSheet.getRange(rowNumber, 1, 1, HEADERS.PROTEIN_LOG.length).setValues([row]);

    return {
      ok: true,
      proteinEntry: parseProteinRow_(row, rowNumber),
      addedFood: resolved.addedFood,
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteProteinEntry(rowNumberInput) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    setupSheets();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.PROTEIN_LOG);
    const rowNumber = parseRowNumber_(rowNumberInput, sheet);
    sheet.deleteRow(rowNumber);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function getDailySnapshot(request) {
  setupSheets();
  const parsed = parseDateProfileRequest_(request);
  const date = parsed.date;
  const profileName = parsed.profileName;
  return {
    date: date,
    profileName: profileName,
    workouts: listWorkoutsByDate_(date, profileName),
    proteins: listProteinEntriesByDate_(date, profileName),
    dailyTotalProtein: sumProteinByDate_(date, profileName),
  };
}

function getWeeklyTrend(request) {
  setupSheets();
  const parsed = parseDateProfileRequest_(request);
  const tz = Session.getScriptTimeZone();
  const endDateLabel = parsed.date;
  const profileName = parsed.profileName;
  const trainingNameFilter = normalizeTrainingNameFilter_(request && request.trainingName);
  const trendDays = normalizeTrendDays_(request && request.trendDays);
  const endDate = parseDateLabel_(endDateLabel);
  const startDate = new Date(endDate);
  startDate.setDate(startDate.getDate() - (trendDays - 1));
  const startLabel = Utilities.formatDate(startDate, tz, "yyyy-MM-dd");

  const proteinByDate = {};
  const proteinRows = listProteinEntriesBetween_(startLabel, endDateLabel, profileName);
  proteinRows.forEach(function (row) {
    if (!proteinByDate[row.date]) {
      proteinByDate[row.date] = 0;
    }
    proteinByDate[row.date] += Number(row.proteinGrams || 0);
  });

  const workoutByDate = {};
  const workoutRows = listWorkoutsBetween_(
    startLabel,
    endDateLabel,
    profileName,
    trainingNameFilter
  );
  workoutRows.forEach(function (row) {
    if (!workoutByDate[row.date]) {
      workoutByDate[row.date] = { load: 0, strengthVolume: 0, cardioMinutes: 0, count: 0 };
    }
    workoutByDate[row.date].load += Number(row.loadScore || 0);
    workoutByDate[row.date].strengthVolume += Number(row.strengthVolume || 0);
    workoutByDate[row.date].cardioMinutes += Number(row.durationMin || 0);
    workoutByDate[row.date].count += 1;
  });

  const out = [];
  for (let i = 0; i < trendDays; i += 1) {
    const d = new Date(startDate);
    d.setDate(startDate.getDate() + i);
    const key = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    out.push({
      date: key,
      totalProteinGrams: round2_(proteinByDate[key] || 0),
      workoutLoad: round2_(workoutByDate[key] ? workoutByDate[key].load : 0),
      workoutVolume: round2_(workoutByDate[key] ? workoutByDate[key].strengthVolume : 0),
      cardioMinutes: round2_(workoutByDate[key] ? workoutByDate[key].cardioMinutes : 0),
      workoutEntries: workoutByDate[key] ? workoutByDate[key].count : 0,
    });
  }

  return out;
}

function normalizeTrendDays_(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) {
    return 7;
  }
  const intVal = Math.floor(n);
  if (intVal <= 7) {
    return 7;
  }
  if (intVal <= 14) {
    return 14;
  }
  return 30;
}

function getProteinTotals(days) {
  setupSheets();
  let dayInput = days;
  let profileName = DEFAULT_PROFILE_NAME;
  if (days && typeof days === "object" && !Array.isArray(days)) {
    dayInput = days.days;
    const profiles = listProfiles_();
    profileName = resolveRequestedProfile_(days.profileName, profiles);
  }

  const n = Number(dayInput);
  const lookbackDays = Number.isFinite(n) && n > 0 ? Math.floor(n) : 14;
  const tz = Session.getScriptTimeZone();
  const end = new Date();
  end.setHours(0, 0, 0, 0);
  const start = new Date(end);
  start.setDate(start.getDate() - (lookbackDays - 1));
  const startLabel = Utilities.formatDate(start, tz, "yyyy-MM-dd");
  const endLabel = Utilities.formatDate(end, tz, "yyyy-MM-dd");

  const proteinMap = {};
  const rows = listProteinEntriesBetween_(startLabel, endLabel, profileName);

  rows.forEach(function (row) {
    if (!proteinMap[row.date]) {
      proteinMap[row.date] = 0;
    }
    proteinMap[row.date] += Number(row.proteinGrams || 0);
  });

  const output = [];
  for (let i = 0; i < lookbackDays; i += 1) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const key = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    output.push({
      date: key,
      totalProteinGrams: round2_(proteinMap[key] || 0),
    });
  }
  return output;
}

function ensureSheet(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = headers.some(function (header, idx) {
    return String(firstRow[idx] || "").trim() !== header;
  });

  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function parseWorkoutPayload_(payload) {
  if (!payload || typeof payload !== "object") {
    throw new Error("Workout payload is required.");
  }

  const profileName = normalizeProfileName_(payload.profileName);
  const workoutType = normalizeWorkoutType_(payload.workoutType);
  const training = String(payload.training || "").trim();
  if (!training) {
    throw new Error("Training name is required.");
  }

  const date = normalizeDateString(payload.date);
  const notes = String(payload.notes || "").trim();

  if (workoutType === WORKOUT_TYPES.CARDIO) {
    const durationMin = parseRequiredPositiveNumber_(payload.durationMin, "Duration");
    const intensity = parseOptionalPositiveNumber_(payload.intensity, "Intensity");
    const speedKph = parseOptionalPositiveNumber_(payload.speedKph, "Speed");

    return [
      new Date(),
      date,
      profileName,
      training,
      WORKOUT_TYPES.CARDIO,
      "",
      "",
      "",
      round2_(durationMin),
      intensity === "" ? "" : round2_(intensity),
      speedKph === "" ? "" : round2_(speedKph),
      notes,
    ];
  }

  const weightKg = parseOptionalNonNegativeNumber_(payload.weightKg, "Weight");
  const reps = parseRequiredPositiveInteger_(payload.reps, "Reps");
  const sets = parseRequiredPositiveInteger_(payload.sets, "Sets");

  return [
    new Date(),
    date,
    profileName,
    training,
    WORKOUT_TYPES.STRENGTH,
    weightKg === "" ? 0 : round2_(weightKg),
    reps,
    sets,
    "",
    "",
    "",
    notes,
  ];
}

function parseFoodPayload_(payload) {
  if (!payload || typeof payload !== "object") {
    throw new Error("Food payload is required.");
  }
  const foodName = normalizeFoodName_(payload.foodName);
  const proteinPer100g = Number(payload.proteinPer100g);

  if (!Number.isFinite(proteinPer100g) || proteinPer100g <= 0 || proteinPer100g > 100) {
    throw new Error("Protein per 100g must be greater than 0 and at most 100.");
  }

  return {
    foodName: foodName,
    foodKey: foodName.toLowerCase(),
    proteinPer100g: round2_(proteinPer100g),
  };
}

function parseProteinPayload_(payload) {
  if (!payload || typeof payload !== "object") {
    throw new Error("Protein entry payload is required.");
  }

  const profileName = normalizeProfileName_(payload.profileName);
  const foodName = normalizeFoodName_(payload.foodName);
  const intakeGrams = Number(payload.intakeGrams);
  const rawProtein = payload.proteinPer100g;
  const proteinPer100g =
    rawProtein === "" || rawProtein === null || typeof rawProtein === "undefined"
      ? 0
      : Number(rawProtein);

  if (!Number.isFinite(intakeGrams) || intakeGrams <= 0) {
    throw new Error("Intake grams must be greater than 0.");
  }

  if (!Number.isFinite(proteinPer100g) || proteinPer100g < 0 || proteinPer100g > 100) {
    throw new Error("Protein per 100g must be between 0 and 100.");
  }

  return {
    date: normalizeDateString(payload.date),
    profileName: profileName,
    foodName: foodName,
    foodKey: foodName.toLowerCase(),
    intakeGrams: round2_(intakeGrams),
    proteinPer100g: round2_(proteinPer100g),
    notes: String(payload.notes || "").trim(),
  };
}

function normalizeFoodName_(value) {
  const foodName = String(value || "").trim();
  if (!foodName) {
    throw new Error("Food name is required.");
  }
  return foodName;
}

function normalizeProfileName_(value) {
  const profileName = String(value || DEFAULT_PROFILE_NAME).trim();
  if (!profileName) {
    throw new Error("Profile name is required.");
  }
  if (profileName.length > 40) {
    throw new Error("Profile name must be 40 characters or fewer.");
  }
  return profileName;
}

function normalizeStoredProfileName_(value) {
  const clean = String(value || "").trim();
  return clean || DEFAULT_PROFILE_NAME;
}

function parseDateProfileRequest_(request) {
  const profiles = listProfiles_();
  if (request && typeof request === "object" && !Array.isArray(request)) {
    return {
      date: normalizeDateString(request.date),
      profileName: resolveRequestedProfile_(request.profileName, profiles),
    };
  }
  return {
    date: normalizeDateString(request),
    profileName: resolveRequestedProfile_("", profiles),
  };
}

function resolveRequestedProfile_(profileInput, profiles) {
  const list = Array.isArray(profiles) && profiles.length ? profiles : [DEFAULT_PROFILE_NAME];
  const requested = String(profileInput || "").trim().toLowerCase();
  if (!requested) {
    return list[0];
  }
  for (let i = 0; i < list.length; i += 1) {
    if (String(list[i]).trim().toLowerCase() === requested) {
      return list[i];
    }
  }
  return list[0];
}

function normalizeWorkoutType_(value) {
  const clean = String(value || WORKOUT_TYPES.STRENGTH).trim().toLowerCase();
  return clean === WORKOUT_TYPES.CARDIO ? WORKOUT_TYPES.CARDIO : WORKOUT_TYPES.STRENGTH;
}

function normalizeTrainingNameFilter_(value) {
  return String(value || "").trim();
}

function resolveProteinPer100g_(clean, foodSheet) {
  const foodLookup = buildFoodLookup_(foodSheet);
  const existingFood = foodLookup[clean.foodKey];

  if (existingFood) {
    return {
      proteinPer100g: Number(existingFood.proteinPer100g),
      addedFood: false,
    };
  }

  if (clean.proteinPer100g > 0) {
    upsertFood_(foodSheet, clean.foodName, clean.foodKey, clean.proteinPer100g);
    return {
      proteinPer100g: clean.proteinPer100g,
      addedFood: true,
    };
  }

  throw new Error(
    "New food detected. Please provide protein per 100g so it can be saved for next time."
  );
}

function upsertFood_(sheet, foodName, foodKey, proteinPer100g) {
  const all = sheet.getDataRange().getValues();

  for (let i = 1; i < all.length; i += 1) {
    if (String(all[i][1]).trim().toLowerCase() === foodKey) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[foodName, round2_(proteinPer100g)]]);
      return true;
    }
  }

  sheet.appendRow([new Date(), foodName, round2_(proteinPer100g)]);
  return false;
}

function normalizeDateString(value) {
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
    return value.trim();
  }

  const date = value ? new Date(value) : new Date();
  if (Object.prototype.toString.call(date) !== "[object Date]" || isNaN(date.getTime())) {
    throw new Error("Invalid date.");
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function round2_(value) {
  return Math.round(Number(value) * 100) / 100;
}

function parseRequiredPositiveNumber_(value, fieldName) {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) {
    throw new Error(fieldName + " must be greater than 0.");
  }
  return n;
}

function parseOptionalPositiveNumber_(value, fieldName) {
  if (value === "" || value === null || typeof value === "undefined") {
    return "";
  }
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) {
    throw new Error(fieldName + " must be greater than 0.");
  }
  return n;
}

function parseOptionalNonNegativeNumber_(value, fieldName) {
  if (value === "" || value === null || typeof value === "undefined") {
    return "";
  }
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) {
    throw new Error(fieldName + " must be 0 or greater.");
  }
  return n;
}

function parseRequiredPositiveInteger_(value, fieldName) {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) {
    throw new Error(fieldName + " must be greater than 0.");
  }
  return Math.floor(n);
}

function parseRowNumber_(value, sheet) {
  const rowNumber = Number(value);
  if (!Number.isInteger(rowNumber) || rowNumber < 2) {
    throw new Error("Invalid row number.");
  }
  if (rowNumber > sheet.getLastRow()) {
    throw new Error("Entry no longer exists. Refresh and try again.");
  }
  return rowNumber;
}

function parseDateLabel_(label) {
  const parts = String(label || "").split("-");
  if (parts.length !== 3) {
    throw new Error("Invalid date label.");
  }
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  const d = Number(parts[2]);
  return new Date(y, m - 1, d);
}

function buildFoodLookup_(sheet) {
  const values = sheet.getDataRange().getValues();
  const lookup = {};
  for (let i = 1; i < values.length; i += 1) {
    const name = String(values[i][1] || "").trim();
    const key = name.toLowerCase();
    const proteinPer100g = Number(values[i][2]);
    if (name && Number.isFinite(proteinPer100g)) {
      lookup[key] = {
        foodName: name,
        proteinPer100g: round2_(proteinPer100g),
      };
    }
  }
  return lookup;
}

function listFoods_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.FOOD_CATALOG);
  const values = sheet.getDataRange().getValues();

  const out = [];
  for (let i = 1; i < values.length; i += 1) {
    const foodName = String(values[i][1] || "").trim();
    const proteinPer100g = Number(values[i][2]);
    if (foodName && Number.isFinite(proteinPer100g)) {
      out.push({
        foodName: foodName,
        proteinPer100g: round2_(proteinPer100g),
      });
    }
  }

  out.sort(function (a, b) {
    return a.foodName.localeCompare(b.foodName);
  });

  return out;
}

function ensureDefaultProfile_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.PROFILES);
  if (!findProfileName_(sheet, DEFAULT_PROFILE_NAME)) {
    sheet.appendRow([new Date(), DEFAULT_PROFILE_NAME]);
  }
}

function ensureProfileExistsByName_(profileName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.PROFILES);
  if (!findProfileName_(sheet, profileName)) {
    sheet.appendRow([new Date(), profileName]);
  }
}

function findProfileName_(sheet, name) {
  const wanted = String(name || "").trim().toLowerCase();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    const candidate = String(values[i][1] || "").trim();
    if (candidate && candidate.toLowerCase() === wanted) {
      return candidate;
    }
  }
  return "";
}

function listProfiles_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.PROFILES);
  const values = sheet.getDataRange().getValues();
  const seen = {};
  const out = [];

  for (let i = 1; i < values.length; i += 1) {
    const profileName = normalizeStoredProfileName_(values[i][1]);
    const key = profileName.toLowerCase();
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    out.push(profileName);
  }

  if (!seen[DEFAULT_PROFILE_NAME.toLowerCase()]) {
    out.unshift(DEFAULT_PROFILE_NAME);
  }

  return out;
}

function listWorkoutNames_(profileName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.WORKOUT_LOG);
  const values = sheet.getDataRange().getValues();
  const out = [];
  const seen = {};

  for (let i = 1; i < values.length; i += 1) {
    const parsed = parseWorkoutRow_(values[i], i + 1);
    if (parsed.profileName !== profileName) {
      continue;
    }
    const training = String(parsed.training || "").trim();
    if (!training) {
      continue;
    }
    const key = normalizeNameKey_(training);
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    out.push(training);
  }

  out.sort(function (a, b) {
    return a.localeCompare(b);
  });
  return out;
}

function listRecentWorkoutNames_(profileName, limitInput) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.WORKOUT_LOG);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  const out = [];
  const seen = {};
  const limit = Math.max(1, Math.min(20, Math.floor(Number(limitInput) || 8)));

  for (let i = 1; i < values.length; i += 1) {
    const parsed = parseWorkoutRow_(values[i], i + 1);
    if (parsed.profileName !== profileName) {
      continue;
    }
    const training = String(parsed.training || "").trim();
    if (!training) {
      continue;
    }
    rows.push({
      training: training,
      key: normalizeNameKey_(training),
      ts: toTimestamp_(parsed.createdAt),
    });
  }

  rows.sort(function (a, b) {
    return b.ts - a.ts;
  });

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (seen[row.key]) {
      continue;
    }
    seen[row.key] = true;
    out.push(row.training);
    if (out.length >= limit) {
      break;
    }
  }

  return out;
}

function listWorkoutsByDate_(date, profileName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.WORKOUT_LOG);
  const values = sheet.getDataRange().getValues();

  const out = [];
  for (let i = 1; i < values.length; i += 1) {
    const parsed = parseWorkoutRow_(values[i], i + 1);
    if (parsed.date !== date || parsed.profileName !== profileName) {
      continue;
    }
    out.push(parsed);
  }

  return out;
}

function listProteinEntriesByDate_(date, profileName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.PROTEIN_LOG);
  const values = sheet.getDataRange().getValues();
  const out = [];

  for (let i = 1; i < values.length; i += 1) {
    const parsed = parseProteinRow_(values[i], i + 1);
    if (parsed.date !== date || parsed.profileName !== profileName) {
      continue;
    }
    out.push(parsed);
  }

  return out;
}

function listProteinEntriesBetween_(startDate, endDate, profileName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.PROTEIN_LOG);
  const values = sheet.getDataRange().getValues();
  const out = [];

  for (let i = 1; i < values.length; i += 1) {
    const parsed = parseProteinRow_(values[i], i + 1);
    if (parsed.date < startDate || parsed.date > endDate || parsed.profileName !== profileName) {
      continue;
    }
    out.push({
      date: parsed.date,
      proteinGrams: Number(parsed.proteinGrams || 0),
    });
  }
  return out;
}

function listWorkoutsBetween_(startDate, endDate, profileName, trainingNameFilter) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.WORKOUT_LOG);
  const values = sheet.getDataRange().getValues();
  const out = [];
  const filterKey = normalizeNameKey_(trainingNameFilter);

  for (let i = 1; i < values.length; i += 1) {
    const parsed = parseWorkoutRow_(values[i], i + 1);
    if (parsed.date < startDate || parsed.date > endDate || parsed.profileName !== profileName) {
      continue;
    }
    if (filterKey && normalizeNameKey_(parsed.training) !== filterKey) {
      continue;
    }
    out.push(parsed);
  }
  return out;
}

function parseWorkoutRow_(row, rowNumber) {
  const dateLabel = coerceDateLabel_(row[1]);
  const typeAt4 = String(row[4] || "").trim().toLowerCase();
  const typeAt3 = String(row[3] || "").trim().toLowerCase();
  const isProfileFormat = typeAt4 === WORKOUT_TYPES.STRENGTH || typeAt4 === WORKOUT_TYPES.CARDIO;
  const isTypedOldFormat =
    !isProfileFormat && (typeAt3 === WORKOUT_TYPES.STRENGTH || typeAt3 === WORKOUT_TYPES.CARDIO);

  let workoutType = WORKOUT_TYPES.STRENGTH;
  let profileRaw = DEFAULT_PROFILE_NAME;
  let trainingRaw = "";
  let weightRaw = "";
  let repsRaw = "";
  let setsRaw = "";
  let durationRaw = "";
  let intensityRaw = "";
  let speedRaw = "";
  let notesRaw = "";

  if (isProfileFormat) {
    profileRaw = row[2];
    trainingRaw = row[3];
    workoutType = typeAt4;
    weightRaw = row[5];
    repsRaw = row[6];
    setsRaw = row[7];
    durationRaw = row[8];
    intensityRaw = row[9];
    speedRaw = row[10];
    notesRaw = row[11];
  } else if (isTypedOldFormat) {
    trainingRaw = row[2];
    workoutType = typeAt3;
    weightRaw = row[4];
    repsRaw = row[5];
    setsRaw = row[6];
    durationRaw = row[7];
    intensityRaw = row[8];
    speedRaw = row[9];
    notesRaw = row[10];
  } else {
    // Backward compatibility with oldest schema before workout type columns were added.
    trainingRaw = row[2];
    weightRaw = row[3];
    repsRaw = row[4];
    setsRaw = row[5];
    notesRaw = row[6];
  }

  const weightKg = toNumberOrNull_(weightRaw);
  const reps = toIntOrNull_(repsRaw);
  const sets = toIntOrNull_(setsRaw);
  const durationMin = toNumberOrNull_(durationRaw);
  const intensity = toNumberOrNull_(intensityRaw);
  const speedKph = toNumberOrNull_(speedRaw);

  if (!isProfileFormat && !isTypedOldFormat) {
    const hasStrengthPattern = reps !== null && sets !== null;
    const hasCardioPattern = durationMin !== null && durationMin > 0;
    workoutType = hasCardioPattern && !hasStrengthPattern ? WORKOUT_TYPES.CARDIO : WORKOUT_TYPES.STRENGTH;
  }

  const strengthVolume =
    workoutType === WORKOUT_TYPES.STRENGTH
      ? round2_((weightKg || 0) * (reps || 0) * (sets || 0))
      : 0;
  const cardioLoad = workoutType === WORKOUT_TYPES.CARDIO ? computeCardioLoad_(durationMin, intensity, speedKph) : 0;

  return {
    rowNumber: rowNumber,
    createdAt: row[0],
    date: dateLabel,
    profileName: normalizeStoredProfileName_(profileRaw),
    training: String(trainingRaw || ""),
    workoutType: workoutType,
    weightKg: weightKg,
    reps: reps,
    sets: sets,
    durationMin: durationMin,
    intensity: intensity,
    speedKph: speedKph,
    notes: String(notesRaw || ""),
    strengthVolume: strengthVolume,
    cardioLoad: cardioLoad,
    loadScore: round2_(strengthVolume + cardioLoad),
  };
}

function parseProteinRow_(row, rowNumber) {
  const dateLabel = coerceDateLabel_(row[1]);
  const intakeAt3 = toNumberOrNull_(row[3]);
  const hasProfileFormat = intakeAt3 === null;

  let profileRaw = DEFAULT_PROFILE_NAME;
  let foodRaw = "";
  let intakeRaw = "";
  let proteinPer100gRaw = "";
  let proteinGramsRaw = "";
  let notesRaw = "";

  if (hasProfileFormat) {
    profileRaw = row[2];
    foodRaw = row[3];
    intakeRaw = row[4];
    proteinPer100gRaw = row[5];
    proteinGramsRaw = row[6];
    notesRaw = row[7];
  } else {
    // Backward compatibility with old schema before profile column was added.
    foodRaw = row[2];
    intakeRaw = row[3];
    proteinPer100gRaw = row[4];
    proteinGramsRaw = row[5];
    notesRaw = row[6];
  }

  return {
    rowNumber: rowNumber,
    createdAt: row[0],
    date: dateLabel,
    profileName: normalizeStoredProfileName_(profileRaw),
    foodName: String(foodRaw || ""),
    intakeGrams: Number(intakeRaw) || 0,
    proteinPer100g: Number(proteinPer100gRaw) || 0,
    proteinGrams: Number(proteinGramsRaw) || 0,
    notes: String(notesRaw || ""),
  };
}

function computeCardioLoad_(durationMin, intensity, speedKph) {
  const duration = Number(durationMin);
  if (!Number.isFinite(duration) || duration <= 0) {
    return 0;
  }

  const intensityFactor = Number(intensity);
  if (Number.isFinite(intensityFactor) && intensityFactor > 0) {
    return round2_(duration * intensityFactor);
  }

  const speedFactor = Number(speedKph);
  if (Number.isFinite(speedFactor) && speedFactor > 0) {
    return round2_(duration * (speedFactor / 6));
  }

  return round2_(duration);
}

function toNumberOrNull_(value) {
  if (value === "" || value === null || typeof value === "undefined") {
    return null;
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function toIntOrNull_(value) {
  const n = toNumberOrNull_(value);
  return n === null ? null : Math.floor(n);
}

function normalizeNameKey_(value) {
  return String(value || "").trim().toLowerCase();
}

function coerceDateLabel_(value) {
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
    return value.trim();
  }
  const date = new Date(value);
  if (Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date.getTime())) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value || "");
}

function sumProteinByDate_(date, profileName) {
  return round2_(
    listProteinEntriesByDate_(date, profileName).reduce(function (sum, row) {
      return sum + Number(row.proteinGrams || 0);
    }, 0)
  );
}
