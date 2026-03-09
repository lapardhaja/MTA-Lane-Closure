const csvInput = document.getElementById("csvInput");
const clearData = document.getElementById("clearData");
const locationFilter = document.getElementById("locationFilter");
const directionFilter = document.getElementById("directionFilter");
const weekdayFilter = document.getElementById("weekdayFilter");
const startDate = document.getElementById("startDate");
const endDate = document.getElementById("endDate");
const holidayUS = document.getElementById("holidayUS");
const holidayME = document.getElementById("holidayME");
const holidayNH = document.getElementById("holidayNH");
const holidayPanel = document.getElementById("holidayPanel");
const toggleHolidays = document.getElementById("toggleHolidays");
const toggleHolidayDetail = document.getElementById("toggleHolidayDetail");
const holidayDetailList = document.getElementById("holidayDetailList");
const resetFilters = document.getElementById("resetFilters");
const exportFormat = document.getElementById("exportFormat");
const exportView = document.getElementById("exportView");
const tableHead = document.getElementById("tableHead");
const tableBody = document.getElementById("tableBody");
const tableWrap = document.getElementById("tableWrap");
const tableScrollbar = document.getElementById("tableScrollbar");
const tableScrollbarInner = document.getElementById("tableScrollbarInner");
const tableMeta = document.getElementById("tableMeta");
const rowCount = document.getElementById("rowCount");
const naSummaryCount = document.getElementById("naSummaryCount");
const naSummaryList = document.getElementById("naSummaryList");
const toggleNaSummary = document.getElementById("toggleNaSummary");
const naSummaryPanel = document.getElementById("naSummaryPanel");
const fileName = document.getElementById("fileName");
const toggleSettings = document.getElementById("toggleSettings");
const settingsPanel = document.getElementById("settingsPanel");
const toggleHolidayMapping = document.getElementById("toggleHolidayMapping");
const projectionHolidayPanel = document.getElementById("projectionHolidayPanel");
const thresholdsPanel = document.getElementById("thresholdsPanel");
const thresholdsMeta = document.getElementById("thresholdsMeta");
const threshold2Single = document.getElementById("threshold-2-single");
const threshold3Double = document.getElementById("threshold-3-double");
const threshold3Single = document.getElementById("threshold-3-single");
const threshold2SingleLabel = document.getElementById("threshold-2-single-label");
const threshold2NoneLabel = document.getElementById("threshold-2-none-label");
const threshold2WhiteLabel = document.getElementById("threshold-2-white-label");
const threshold2YellowLabel = document.getElementById("threshold-2-yellow-label");
const threshold2OrangeLabel = document.getElementById("threshold-2-orange-label");
const threshold3DoubleLabel = document.getElementById("threshold-3-double-label");
const threshold3SingleLabel = document.getElementById("threshold-3-single-label");
const threshold3NoneLabel = document.getElementById("threshold-3-none-label");
const threshold3WhiteLabel = document.getElementById("threshold-3-white-label");
const threshold3YellowLabel = document.getElementById("threshold-3-yellow-label");
const threshold3Error = document.getElementById("threshold-3-error");
const tabCurrent = document.getElementById("tabCurrent");
const tabProjections = document.getElementById("tabProjections");
const currentView = document.getElementById("currentView");
const projectionsView = document.getElementById("projectionsView");
const projectionYear = document.getElementById("projectionYear");
const projectionLocationRates = document.getElementById("projectionLocationRates");
const projectionHolidaySummary = document.getElementById(
  "projectionHolidaySummary"
);
const projectionTableHead = document.getElementById("projectionTableHead");
const projectionTableBody = document.getElementById("projectionTableBody");
const projectionTableWrap = document.getElementById("projectionTableWrap");
const projectionTableScrollbar = document.getElementById(
  "projectionTableScrollbar"
);
const projectionTableScrollbarInner = document.getElementById(
  "projectionTableScrollbarInner"
);
const projectionTableMeta = document.getElementById("projectionTableMeta");
const closureScheduleToggles = document.getElementById("closureScheduleToggles");
const closureToggleSingle = document.getElementById("closureToggleSingle");
const closureToggleDouble = document.getElementById("closureToggleDouble");
const closureScheduleMeta = document.getElementById("closureScheduleMeta");
const closureScheduleHint = document.getElementById("closureScheduleHint");
const closureScheduleBeginHead = document.getElementById("closureScheduleBeginHead");
const closureScheduleBeginBody = document.getElementById("closureScheduleBeginBody");
const closureScheduleEndHead = document.getElementById("closureScheduleEndHead");
const closureScheduleEndBody = document.getElementById("closureScheduleEndBody");
const closureScheduleTablesWrap = document.getElementById("closureScheduleTablesWrap");
const closureScheduleLegend = document.getElementById("closureScheduleLegend");
const projectionClosureToggles = document.getElementById("projectionClosureToggles");
const projectionClosureToggleSingle = document.getElementById("projectionClosureToggleSingle");
const projectionClosureToggleDouble = document.getElementById("projectionClosureToggleDouble");
const projectionClosureScheduleMeta = document.getElementById("projectionClosureScheduleMeta");
const projectionClosureScheduleHint = document.getElementById("projectionClosureScheduleHint");
const projectionClosureScheduleBeginHead = document.getElementById("projectionClosureScheduleBeginHead");
const projectionClosureScheduleBeginBody = document.getElementById("projectionClosureScheduleBeginBody");
const projectionClosureScheduleEndHead = document.getElementById("projectionClosureScheduleEndHead");
const projectionClosureScheduleEndBody = document.getElementById("projectionClosureScheduleEndBody");
const projectionClosureScheduleTablesWrap = document.getElementById("projectionClosureScheduleTablesWrap");
const projectionClosureScheduleLegend = document.getElementById("projectionClosureScheduleLegend");

const HOLIDAYS_BY_REGION = {
  us: [
    "New Year's Day",
    "Martin Luther King Jr. Day",
    "Washington's Birthday",
    "Memorial Day",
    "Juneteenth",
    "Independence Day",
    "Labor Day",
    "Columbus Day",
    "Veterans Day",
    "Thanksgiving Day",
    "Christmas Day",
  ],
  me: ["Patriots' Day"],
  nh: ["Civil Rights Day", "Patriots' Day"],
};

const state = {
  rows: [],
  filtered: [],
  thresholds: {
    2: { singleMax: 1340 },
    3: { doubleMax: 1350, singleMax: 3000 },
  },
  projection: {
    targetYear: null,
    defaultRate: 2,
    locationRates: {},
  },
  holidays: {
    disabledHolidays: new Set(),
  },
  closureScheduleType: "single",
};

let thresholdApplyTimer = null;

let syncingScroll = false;
function syncScroll(source, target) {
  if (syncingScroll) return;
  syncingScroll = true;
  target.scrollLeft = source.scrollLeft;
  syncingScroll = false;
}

if (tableWrap && tableScrollbar) {
  tableWrap.addEventListener("scroll", () =>
    syncScroll(tableWrap, tableScrollbar)
  );
  tableScrollbar.addEventListener("scroll", () =>
    syncScroll(tableScrollbar, tableWrap)
  );
}

if (projectionTableWrap && projectionTableScrollbar) {
  projectionTableWrap.addEventListener("scroll", () =>
    syncScroll(projectionTableWrap, projectionTableScrollbar)
  );
  projectionTableScrollbar.addEventListener("scroll", () =>
    syncScroll(projectionTableScrollbar, projectionTableWrap)
  );
}

const headerMap = {
  location: ["location", "loc", "segment"],
  direction: ["direction", "dir"],
  date: ["date", "day"],
  time: ["time", "hour"],
  datetime: ["datetime", "timestamp"],
  volume: ["volume", "vol", "count"],
  lanes: ["lanes", "lane", "lane count", "lanecount", "number of lanes"],
};

function normalizeHeader(name) {
  const key = name.trim().toLowerCase();
  return key.replace(/[^a-z0-9]/g, "");
}

function resolveHeaderIndex(headers, field) {
  const options = headerMap[field];
  for (const option of options) {
    const target = normalizeHeader(option);
    const idx = headers.findIndex((h) => normalizeHeader(h) === target);
    if (idx !== -1) return idx;
  }
  return -1;
}

function splitCSVLine(line) {
  const cells = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      const next = line[i + 1];
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (char === "," && !inQuotes) {
      cells.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }
  cells.push(current.trim());
  return cells;
}

function parseDateSafe(dateStr) {
  if (!dateStr) return null;
  const iso = new Date(dateStr);
  if (!Number.isNaN(iso.getTime())) return iso;
  const cleaned = dateStr.replace(/-/g, "/");
  const fallback = new Date(cleaned);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function parseDateTime(dateStr, timeStr) {
  if (!dateStr) return null;
  if (timeStr) {
    const joined = `${dateStr} ${timeStr}`;
    const candidate = new Date(joined);
    if (!Number.isNaN(candidate.getTime())) return candidate;
  }
  return parseDateSafe(dateStr);
}

function formatDateKey(dateObj, fallback) {
  if (!dateObj) return fallback || "Unknown";
  const year = dateObj.getFullYear();
  const month = String(dateObj.getMonth() + 1).padStart(2, "0");
  const day = String(dateObj.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDateDisplay(dateKey) {
  if (!dateKey || dateKey === "Unknown") return dateKey || "Unknown";
  const parts = dateKey.split("-");
  if (parts.length !== 3) return dateKey;
  const [year, month, day] = parts;
  return `${month}/${day}/${year}`;
}

function mapDateToYearByWeekday(dateObj, targetYear) {
  const sameCalendar = new Date(
    targetYear,
    dateObj.getMonth(),
    dateObj.getDate()
  );
  const baseWeekday = dateObj.getDay();
  const targetWeekday = sameCalendar.getDay();
  let diff = baseWeekday - targetWeekday;
  if (diff > 3) diff -= 7;
  if (diff < -3) diff += 7;
  const adjusted = new Date(sameCalendar);
  adjusted.setDate(adjusted.getDate() + diff);
  if (adjusted.getFullYear() < targetYear) {
    adjusted.setDate(adjusted.getDate() + 7);
  } else if (adjusted.getFullYear() > targetYear) {
    adjusted.setDate(adjusted.getDate() - 7);
  }
  return adjusted;
}

function getBaseYear(rows) {
  const counts = new Map();
  for (const row of rows) {
    if (!row.dateObj) continue;
    const year = row.dateObj.getFullYear();
    counts.set(year, (counts.get(year) || 0) + 1);
  }
  let bestYear = null;
  let bestCount = 0;
  for (const [year, count] of counts.entries()) {
    if (count > bestCount) {
      bestYear = year;
      bestCount = count;
    }
  }
  return bestYear;
}

function getFullYearDateKeys(targetYear) {
  const dates = [];
  const date = new Date(targetYear, 0, 1);
  while (date.getFullYear() === targetYear) {
    dates.push(formatDateKey(date));
    date.setDate(date.getDate() + 1);
  }
  return dates;
}

function deriveTier(volume, lanes) {
  if (Number.isNaN(volume)) return "unknown";
  const laneCount = Number(lanes);
  if (!Number.isFinite(laneCount)) return "unknown";
  const config = state.thresholds[laneCount];
  if (!config) return "unknown";
  if (laneCount === 2) {
    if (volume > config.singleMax) return "none";
    const pct = config.singleMax > 0 ? volume / config.singleMax : 0;
    if (pct > 0.95) return "single-orange";
    if (pct > 0.9) return "single-yellow";
    return "single-ok";
  }
  if (!Number.isFinite(config.doubleMax) || !Number.isFinite(config.singleMax)) {
    return "unknown";
  }
  if (config.singleMax <= config.doubleMax) {
    return "unknown";
  }
  if (volume > config.singleMax) return "none";
  if (volume > config.doubleMax) return "single-orange";
  const whiteAt = 0.9 * config.doubleMax;
  if (volume <= whiteAt) return "double-ok";
  return "double-yellow";
}

function parseCSV(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (lines.length === 0) return [];

  const headers = splitCSVLine(lines[0]);
  const idxLocation = resolveHeaderIndex(headers, "location");
  const idxDirection = resolveHeaderIndex(headers, "direction");
  const idxDate = resolveHeaderIndex(headers, "date");
  const idxTime = resolveHeaderIndex(headers, "time");
  const idxDateTime = resolveHeaderIndex(headers, "datetime");
  const idxVolume = resolveHeaderIndex(headers, "volume");
  const idxLanes = resolveHeaderIndex(headers, "lanes");

  const rows = [];

  for (let i = 1; i < lines.length; i += 1) {
    const cells = splitCSVLine(lines[i]);
    const location = idxLocation >= 0 ? cells[idxLocation] : "";
    const direction = idxDirection >= 0 ? cells[idxDirection] : "";
    const date = idxDate >= 0 ? cells[idxDate] : "";
    const time = idxTime >= 0 ? cells[idxTime] : "";
    const datetime = idxDateTime >= 0 ? cells[idxDateTime] : "";
    const volumeRaw = idxVolume >= 0 ? cells[idxVolume] : "";
    const volumeClean = volumeRaw.replace(/[^0-9.]/g, "");
    const volume = volumeClean ? Number(volumeClean) : Number.NaN;
    const lanesRaw = idxLanes >= 0 ? cells[idxLanes] : "";
    const lanesClean = lanesRaw.replace(/[^0-9.]/g, "");
    const lanes = lanesClean ? Number(lanesClean) : Number.NaN;

    const resolvedDate = datetime || date;
    const dateObj = parseDateTime(resolvedDate, time);
    const dateKey = formatDateKey(dateObj, resolvedDate);
    const row = {
      location: location || "Unknown",
      direction: (direction || "Unknown").toUpperCase(),
      date: resolvedDate || "Unknown",
      dateKey,
      time: time || "Unknown",
      volume,
      lanes: Number.isFinite(lanes) ? lanes : null,
      closureTier: deriveTier(volume, lanes),
      dateObj,
    };
    rows.push(row);
  }

  return rows;
}

function parseLocationRange(loc) {
  const m = String(loc).match(/(\d+)-(\d+)/);
  return m ? [Number(m[1]), Number(m[2])] : [0, 0];
}

function setLocations(rows) {
  const locations = Array.from(
    new Set(rows.map((row) => row.location).filter(Boolean))
  ).sort((a, b) => {
    const [a1, a2] = parseLocationRange(a);
    const [b1, b2] = parseLocationRange(b);
    return a1 - b1 || a2 - b2;
  });

  locationFilter.innerHTML = `<option value="">Select</option>${locations
    .map((loc) => `<option value="${loc}">${loc}</option>`)
    .join("")}`;
}

function setDirections(rows, location) {
  const filtered = location ? rows.filter((r) => r.location === location) : rows;
  const directions = Array.from(
    new Set(filtered.map((row) => row.direction).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const currentDir = directionFilter.value;
  directionFilter.innerHTML =
    `<option value="">Select</option>` +
    directions.map((dir) => `<option value="${dir}">${dir}</option>`).join("");
  if (currentDir && directions.includes(currentDir)) {
    directionFilter.value = currentDir;
  }
}

function syncRegionHolidaysToParent(region) {
  const labels = HOLIDAYS_BY_REGION[region];
  if (!labels) return;
  const parentChecked =
    (region === "us" && holidayUS && holidayUS.checked) ||
    (region === "me" && holidayME && holidayME.checked) ||
    (region === "nh" && holidayNH && holidayNH.checked);
  if (!parentChecked) {
    labels.forEach((l) => state.holidays.disabledHolidays.add(l));
  } else {
    labels.forEach((l) => state.holidays.disabledHolidays.delete(l));
  }
}

function renderHolidayDetailList() {
  if (!holidayDetailList) return;
  holidayDetailList.innerHTML = "";
  const regionLabels = { us: "US Federal", me: "Maine", nh: "New Hampshire" };
  const fragment = document.createDocumentFragment();
  for (const [region, labels] of Object.entries(HOLIDAYS_BY_REGION)) {
    const seen = new Set();
    const uniqueLabels = labels.filter((l) => {
      if (seen.has(l)) return false;
      seen.add(l);
      return true;
    });
    const group = document.createElement("div");
    group.className = "holiday-detail-group";
    const heading = document.createElement("div");
    heading.className = "holiday-detail-group-title";
    heading.textContent = regionLabels[region];
    group.appendChild(heading);
    for (const label of uniqueLabels) {
      const labelEl = document.createElement("label");
      labelEl.className = "checkbox holiday-detail-checkbox";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.dataset.holidayLabel = label;
      cb.checked = !state.holidays.disabledHolidays.has(label);
      cb.addEventListener("change", () => {
        if (cb.checked) {
          state.holidays.disabledHolidays.delete(label);
        } else {
          state.holidays.disabledHolidays.add(label);
        }
        renderTable();
        renderProjectionTable();
      });
      labelEl.appendChild(cb);
      labelEl.appendChild(document.createTextNode(" " + label));
      group.appendChild(labelEl);
    }
    fragment.appendChild(group);
  }
  holidayDetailList.appendChild(fragment);
}

function updateRowTiers() {
  for (const row of state.rows) {
    row.closureTier = deriveTier(row.volume, row.lanes);
  }
}

function formatNumber(value) {
  return Number.isFinite(value) ? value.toLocaleString() : "N/A";
}

function readNumberInput(input) {
  if (!input) return Number.NaN;
  const raw = String(input.value || "").trim();
  if (!raw) return Number.NaN;
  const cleaned = raw.replace(/[^0-9]/g, "");
  if (!cleaned) return Number.NaN;
  return Number(cleaned);
}

function formatNumberInputValue(value) {
  const digits = String(value || "").replace(/[^0-9]/g, "");
  if (!digits) return "";
  return Number(digits).toLocaleString();
}

function applyNumericValue(input, value) {
  const formatted = formatNumberInputValue(value);
  input.value = formatted;
  const end = formatted.length;
  input.setSelectionRange(end, end);
}

function readDecimalInput(input) {
  if (!input) return Number.NaN;
  const raw = String(input.value || "").trim();
  if (!raw) return Number.NaN;
  const cleaned = raw.replace(/[^0-9.]/g, "");
  if (!cleaned) return Number.NaN;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : Number.NaN;
}

function formatDecimalInputValue(value) {
  if (value === "" || value == null) return "";
  const n = Number(value);
  if (!Number.isFinite(n)) return "";
  const s = String(n);
  return s;
}

function applyDecimalValue(input, value) {
  const formatted =
    typeof value === "string"
      ? sanitizeDecimalString(value)
      : formatDecimalInputValue(value);
  input.value = formatted;
  const end = formatted.length;
  input.setSelectionRange(end, end);
}

function sanitizeDecimalString(str) {
  const allowed = str.replace(/[^0-9.]/g, "");
  const parts = allowed.split(".");
  if (parts.length <= 1) return allowed;
  return parts[0] + "." + parts.slice(1).join("");
}

function bindDecimalCaretFix(input, onValueChange) {
  if (!input) return;
  input.addEventListener("beforeinput", (event) => {
    if (event.isComposing) return;
    const current = String(input.value || "");
    const start =
      typeof input.selectionStart === "number"
        ? input.selectionStart
        : current.length;
    const end =
      typeof input.selectionEnd === "number" ? input.selectionEnd : start;
    let next = null;
    switch (event.inputType) {
      case "insertText":
      case "insertFromPaste":
      case "insertReplacementText":
      case "insertCompositionText":
        next = current.slice(0, start) + (event.data || "") + current.slice(end);
        break;
      case "deleteContentBackward":
        if (start === end && start > 0) {
          next = current.slice(0, start - 1) + current.slice(end);
        } else {
          next = current.slice(0, start) + current.slice(end);
        }
        break;
      case "deleteContentForward":
        if (start === end && start < current.length) {
          next = current.slice(0, start) + current.slice(end + 1);
        } else {
          next = current.slice(0, start) + current.slice(end);
        }
        break;
      case "deleteByCut":
      case "deleteContent":
        next = current.slice(0, start) + current.slice(end);
        break;
      default:
        return;
    }
    if (typeof next === "string") {
      next = sanitizeDecimalString(next);
    }
    event.preventDefault();
    applyDecimalValue(input, next);
    if (typeof onValueChange === "function") {
      onValueChange();
    }
  });
  input.addEventListener("focus", () => {
    applyDecimalValue(input, input.value);
  });
  input.addEventListener("input", () => {
    applyDecimalValue(input, input.value);
  });
}

function bindNumericCaretFix(input, onValueChange) {
  if (!input) return;
  input.addEventListener("beforeinput", (event) => {
    if (event.isComposing) return;
    const current = String(input.value || "");
    const start =
      typeof input.selectionStart === "number"
        ? input.selectionStart
        : current.length;
    const end =
      typeof input.selectionEnd === "number" ? input.selectionEnd : start;
    let next = null;

    switch (event.inputType) {
      case "insertText":
      case "insertFromPaste":
      case "insertReplacementText":
      case "insertCompositionText":
        next = current.slice(0, start) + (event.data || "") + current.slice(end);
        break;
      case "deleteContentBackward":
        if (start === end && start > 0) {
          next = current.slice(0, start - 1) + current.slice(end);
        } else {
          next = current.slice(0, start) + current.slice(end);
        }
        break;
      case "deleteContentForward":
        if (start === end && start < current.length) {
          next = current.slice(0, start) + current.slice(end + 1);
        } else {
          next = current.slice(0, start) + current.slice(end);
        }
        break;
      case "deleteByCut":
      case "deleteContent":
        next = current.slice(0, start) + current.slice(end);
        break;
      default:
        return;
    }

    if (typeof next === "string") {
      next = next.replace(/[^0-9]/g, "");
    }

    event.preventDefault();
    applyNumericValue(input, next);
    if (typeof onValueChange === "function") {
      onValueChange();
    }
  });

  input.addEventListener("focus", () => {
    applyNumericValue(input, input.value);
  });

  input.addEventListener("input", () => {
    applyNumericValue(input, input.value);
  });
}

function refreshThresholdLabels(values = null) {
  const two = values?.two ?? state.thresholds[2];
  const three = values?.three ?? state.thresholds[3];
  if (threshold2SingleLabel) {
    threshold2SingleLabel.textContent = `volume <= ${formatNumber(two.singleMax)}`;
  }
  if (threshold2NoneLabel) {
    threshold2NoneLabel.textContent = `> ${formatNumber(two.singleMax)}`;
  }
  const t90 = 0.9 * (two.singleMax || 0);
  const t95 = 0.95 * (two.singleMax || 0);
  if (threshold2WhiteLabel) {
    threshold2WhiteLabel.textContent = formatNumber(Math.round(t90));
  }
  if (threshold2YellowLabel) {
    threshold2YellowLabel.textContent = `${formatNumber(Math.round(t90))}–${formatNumber(Math.round(t95))}`;
  }
  if (threshold2OrangeLabel) {
    threshold2OrangeLabel.textContent = `${formatNumber(Math.round(t95))}–${formatNumber(two.singleMax)}`;
  }
  if (threshold3DoubleLabel) {
    threshold3DoubleLabel.textContent = `volume <= ${formatNumber(three.doubleMax)}`;
  }
  if (threshold3SingleLabel) {
    threshold3SingleLabel.textContent = `${formatNumber(three.doubleMax)}–${formatNumber(three.singleMax)}`;
  }
  if (threshold3NoneLabel) {
    threshold3NoneLabel.textContent = `> ${formatNumber(three.singleMax)}`;
  }
  const whiteAt = 0.9 * (three.doubleMax || 0);
  if (threshold3WhiteLabel) {
    threshold3WhiteLabel.textContent = formatNumber(Math.round(whiteAt));
  }
  if (threshold3YellowLabel) {
    threshold3YellowLabel.textContent = `${formatNumber(Math.round(whiteAt))}–${formatNumber(three.doubleMax)}`;
  }

  if (threshold3Error) {
    if (Number.isFinite(three.doubleMax) && Number.isFinite(three.singleMax) && three.singleMax <= three.doubleMax) {
      threshold3Error.textContent =
        "Single closure max must be larger than double closure max.";
    } else {
      threshold3Error.textContent = "";
    }
  }
}

function refreshThresholdLabelsFromInputs() {
  const twoSingle = readNumberInput(threshold2Single);
  const threeDouble = readNumberInput(threshold3Double);
  const threeSingle = readNumberInput(threshold3Single);
  const values = {
    two: {
      singleMax: Number.isFinite(twoSingle)
        ? twoSingle
        : state.thresholds[2].singleMax,
    },
    three: {
      doubleMax: Number.isFinite(threeDouble)
        ? threeDouble
        : state.thresholds[3].doubleMax,
      singleMax: Number.isFinite(threeSingle)
        ? threeSingle
        : state.thresholds[3].singleMax,
    },
  };
  if (values.three.doubleMax > values.three.singleMax) {
    // keep user input; show error instead of adjusting values
  }
  refreshThresholdLabels(values);
}

function updateThresholdInputs() {
  if (threshold2Single) {
    applyNumericValue(threshold2Single, state.thresholds[2].singleMax);
  }
  if (threshold3Double) {
    applyNumericValue(threshold3Double, state.thresholds[3].doubleMax);
  }
  if (threshold3Single) {
    applyNumericValue(threshold3Single, state.thresholds[3].singleMax);
  }
  refreshThresholdLabels();
}

function updateThresholdsFromInputs() {
  const twoSingle = readNumberInput(threshold2Single);
  const threeDouble = readNumberInput(threshold3Double);
  const threeSingle = readNumberInput(threshold3Single);

  if (Number.isFinite(twoSingle)) {
    state.thresholds[2].singleMax = twoSingle;
  }
  if (Number.isFinite(threeDouble)) {
    state.thresholds[3].doubleMax = threeDouble;
  }
  if (Number.isFinite(threeSingle)) {
    state.thresholds[3].singleMax = threeSingle;
  }

  updateThresholdInputs();
  updateRowTiers();
  applyFilters();
  renderTable();
  renderProjectionTable();
  renderClosureSchedule();
  renderProjectionClosureSchedule();
}

function scheduleThresholdApply() {
  if (thresholdApplyTimer) {
    clearTimeout(thresholdApplyTimer);
  }
  thresholdApplyTimer = setTimeout(() => {
    updateThresholdsFromInputs();
    thresholdApplyTimer = null;
  }, 150);
}

function getLaneCountsForSelection() {
  if (state.rows.length === 0) return [];
  const selection =
    locationFilter.value === "all" ? state.rows : state.rows.filter((row) => row.location === locationFilter.value);
  const counts = new Set();
  for (const row of selection) {
    if (Number.isFinite(row.lanes)) counts.add(row.lanes);
  }
  return Array.from(counts).sort((a, b) => a - b);
}

function updateThresholdPanel() {
  const groups = Array.from(document.querySelectorAll("[data-lanes]"));
  if (state.rows.length === 0) {
    groups.forEach((group) => (group.style.display = ""));
    if (thresholdsMeta) {
      thresholdsMeta.textContent =
        "Defaults shown. Load data to auto-select lane controls.";
    }
    return;
  }

  const laneCounts = getLaneCountsForSelection();
  if (laneCounts.length === 0) {
    groups.forEach((group) => (group.style.display = ""));
    if (thresholdsMeta) {
      thresholdsMeta.textContent =
        "No lane count values found for this selection.";
    }
    return;
  }

  const supported = new Set([2, 3]);
  const supportedCounts = laneCounts.filter((lane) => supported.has(lane));
  groups.forEach((group) => {
    const lane = Number(group.dataset.lanes);
    group.style.display = supportedCounts.includes(lane) ? "" : "none";
  });

  const unsupported = laneCounts.filter((lane) => !supported.has(lane));
  if (thresholdsMeta) {
    thresholdsMeta.textContent = unsupported.length
      ? `Lane counts in selection: ${laneCounts.join(
          ", "
        )}. No controls for ${unsupported.join(", ")} lanes.`
      : `Lane counts in selection: ${laneCounts.join(", ")}.`;
  }
}

function getFilters() {
  return {
    location: locationFilter.value,
    direction: directionFilter.value,
    weekday: weekdayFilter.value,
    start: startDate.value ? new Date(startDate.value) : null,
    end: endDate.value ? new Date(endDate.value) : null,
  };
}

function applyFilters() {
  const filters = getFilters();
  state.filtered = state.rows.filter((row) => {
    if (!filters.location || !filters.direction) {
      return false;
    }
    if (row.location !== filters.location) {
      return false;
    }
    if (row.direction !== filters.direction) {
      return false;
    }
    if (filters.weekday !== "all" && row.dateObj) {
      if (filters.weekday === "mon-thu") {
        const d = row.dateObj.getDay();
        if (d < 1 || d > 4) return false; // Mon=1..Thu=4
      } else if (filters.weekday === "mon-wed") {
        const d = row.dateObj.getDay();
        if (d < 1 || d > 3) return false; // Mon=1..Wed=3
      } else {
        const weekday = String(row.dateObj.getDay());
        if (weekday !== filters.weekday) return false;
      }
    }
    if (filters.start && row.dateObj && row.dateObj < filters.start) {
      return false;
    }
    if (filters.end && row.dateObj) {
      const end = new Date(filters.end);
      end.setHours(23, 59, 59, 999);
      if (row.dateObj > end) return false;
    }
    return true;
  });
}

/** Rows for projection: same as filtered but without weekday filter (all days). */
function getRowsForProjection(targetYear) {
  const filters = getFilters();
  if (!filters.location || !filters.direction) return [];
  const yearStart = targetYear != null ? new Date(targetYear, 0, 1) : null;
  const yearEnd = targetYear != null ? new Date(targetYear, 11, 31, 23, 59, 59, 999) : null;
  const filterOverlapsTarget =
    targetYear != null &&
    (filters.start || filters.end) &&
    ((!filters.start || filters.start <= yearEnd) && (!filters.end || filters.end >= yearStart));
  if (filterOverlapsTarget) {
    return state.rows.filter((row) => {
      if (row.location !== filters.location) return false;
      if (row.direction !== filters.direction) return false;
      return true;
    });
  }
  return state.rows.filter((row) => {
    if (row.location !== filters.location) return false;
    if (row.direction !== filters.direction) return false;
    if (filters.start && row.dateObj && row.dateObj < filters.start) return false;
    if (filters.end && row.dateObj) {
      const end = new Date(filters.end);
      end.setHours(23, 59, 59, 999);
      if (row.dateObj > end) return false;
    }
    return true;
  });
}

function renderTable() {
  tableBody.innerHTML = "";
  tableHead.innerHTML = "";

  if (state.rows.length === 0) {
    tableMeta.textContent = "No data loaded.";
    rowCount.textContent = "0";
    if (naSummaryCount) naSummaryCount.textContent = "0 N/A cells";
    if (naSummaryList) naSummaryList.textContent = "No N/A values.";
    tableScrollbar.style.display = "none";
    tableScrollbarInner.style.width = "0px";
    return;
  }

  if (state.filtered.length === 0) {
    const filters = getFilters();
    tableMeta.textContent =
      !filters.location || !filters.direction
        ? "Select location and direction."
        : "No rows match the current filters.";
    rowCount.textContent = "0";
    if (tableWrap) tableWrap.dataset.lanes = "";
    if (naSummaryCount) naSummaryCount.textContent = "0 N/A cells";
    if (naSummaryList) naSummaryList.textContent = "No N/A values.";
    tableScrollbar.style.display = "none";
    tableScrollbarInner.style.width = "0px";
    return;
  }

  const filters = getFilters();
  const isMonThu = filters.weekday === "mon-thu";
  const isMonWed = filters.weekday === "mon-wed";
  const isWeekAggregated = isMonThu || isMonWed;
  const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  let dates;
  let aggregated;
  let aggregatedMeta;

  if (isWeekAggregated) {
    dates = Array.from(
      new Set(state.filtered.map((row) => getWeekStartSunday(row.dateKey)))
    ).sort((a, b) => a.localeCompare(b));
    aggregated = new Map();
    aggregatedMeta = new Map();
    for (const row of state.filtered) {
      const weekStart = getWeekStartSunday(row.dateKey);
      const key = `${weekStart}__${row.time}`;
      const vol = Number.isFinite(row.volume) ? row.volume : NaN;
      const existing = aggregated.get(key);
      const existingVol = existing && Number.isFinite(existing.volume) ? existing.volume : -Infinity;
      if (!Number.isNaN(vol) && vol > existingVol) {
        aggregated.set(key, row);
        aggregatedMeta.set(key, {
          dateKey: row.dateKey,
          dayName: DAY_NAMES[row.dateObj ? row.dateObj.getDay() : 0],
          volume: row.volume,
        });
      } else if (!existing) {
        aggregated.set(key, row);
        aggregatedMeta.set(key, {
          dateKey: row.dateKey,
          dayName: DAY_NAMES[row.dateObj ? row.dateObj.getDay() : 0],
          volume: row.volume,
        });
      }
    }
  } else {
    dates = Array.from(
      new Set(state.filtered.map((row) => row.dateKey))
    ).sort((a, b) => a.localeCompare(b));
    aggregated = new Map();
    for (const row of state.filtered) {
      const key = `${row.dateKey}__${row.time}`;
      const existing = aggregated.get(key);
      if (!existing || row.volume > existing.volume) {
        aggregated.set(key, row);
      }
    }
  }

  const times = Array.from(new Set(state.filtered.map((row) => row.time))).sort(
    (a, b) => timeSortKey(a) - timeSortKey(b)
  );

  const yearSet = new Set(
    dates.map((date) => Number(date.split("-")[0])).filter(Boolean)
  );
  const holidayMap = buildHolidayMap(yearSet, getHolidaySelections());

  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthSpans = [];
  for (let i = 0; i < dates.length; ) {
    const [y, m] = dates[i].split("-").map(Number);
    const monthKey = `${y}-${String(m).padStart(2, "0")}`;
    const startIdx = i;
    let count = 0;
    while (i < dates.length && dates[i].slice(0, 7) === monthKey) {
      count++;
      i++;
    }
    const spanDates = dates.slice(startIdx, startIdx + count);
    let dateKeysForClick = spanDates;
    if (isWeekAggregated) {
      dateKeysForClick = [];
      const dayCount = isMonThu ? 4 : 3;
      for (const weekStart of spanDates) {
        const [yw, mw, dw] = weekStart.split("-").map(Number);
        for (let j = 1; j <= dayCount; j++) {
          const d = new Date(yw, mw - 1, dw + j);
          dateKeysForClick.push(formatDateKey(d));
        }
      }
    }
    monthSpans.push({ name: MONTH_NAMES[m - 1], count, dates: dateKeysForClick });
  }
  const monthRow = document.createElement("tr");
  monthRow.innerHTML = `<th></th>${monthSpans.map((s, idx) => {
    const dateKeys = s.dates.length ? ` data-month-dates="${s.dates.join(",")}"` : "";
    const monthAttr = isMonThu ? ` data-month-mon-thu="true"` : isMonWed ? ` data-month-mon-wed="true"` : "";
    return `<th colspan="${s.count}" class="month-label month-label-clickable${idx > 0 ? " month-boundary" : ""}"${dateKeys}${monthAttr}>${s.name}</th>`;
  }).join("")}`;
  tableHead.appendChild(monthRow);

  const headRow = document.createElement("tr");
  const dateHeaders = dates
    .map((date, i) => {
      const holidays = holidayMap.get(date) || [];
      const title = holidays.length > 0 ? holidays.join(", ") : "";
      const cls = holidays.length > 0 ? "holiday-header" : "";
      const monthBoundary = i > 0 && dates[i - 1].slice(0, 7) !== date.slice(0, 7);
      const boundaryCls = monthBoundary ? " month-boundary" : "";
      const displayDate = isWeekAggregated ? `Week of ${formatDateDisplay(date)}` : formatDateDisplay(date);
      const [yw, mw, dw] = date.split("-").map(Number);
      const weekDates = isWeekAggregated
        ? Array.from({ length: isMonThu ? 4 : 3 }, (_, j) => formatDateKey(new Date(yw, mw - 1, dw + j + 1)))
        : null;
      const weekDatesAttr = weekDates ? ` data-week-dates="${weekDates.join(",")}"` : "";
      return `<th class="${cls}${boundaryCls} date-clickable" data-date="${date}"${weekDatesAttr} title="${title}">${displayDate}</th>`;
    })
    .join("");
  headRow.innerHTML = `<th>Time</th>${dateHeaders}`;
  tableHead.appendChild(headRow);

  const fragment = document.createDocumentFragment();
  const naItems = [];

  for (const time of times) {
    const tr = document.createElement("tr");
    const cells = [`<td>${time}</td>`];
    for (let i = 0; i < dates.length; i++) {
      const date = dates[i];
      const key = `${date}__${time}`;
      const row = aggregated.get(key);
      const meta = isWeekAggregated ? aggregatedMeta.get(key) : null;
      const monthBoundary = i > 0 && dates[i - 1].slice(0, 7) !== date.slice(0, 7);
      const boundaryCls = monthBoundary ? " month-boundary" : "";
      const weekdayRangeAttr = isMonWed ? ` data-weekday-range="mon-wed"` : "";
      const tooltipAttrs = isWeekAggregated && meta
        ? ` data-max-date="${meta.dateKey}" data-max-day="${(meta.dayName || "").replace(/"/g, "&quot;")}" data-max-volume="${Number.isFinite(meta.volume) ? meta.volume : ""}"${weekdayRangeAttr}`
        : "";
      if (!row) {
        cells.push(`<td class="cell-unknown${boundaryCls}"${tooltipAttrs}>—</td>`);
        continue;
      }
      if (Number.isNaN(row.volume)) {
        cells.push(`<td class="cell-na${boundaryCls}"${tooltipAttrs}>N/A</td>`);
        naItems.push({ date: isWeekAggregated ? meta?.dateKey : date, time });
        continue;
      }
      const volume = formatNumber(row.volume);
      cells.push(`<td class="cell-${row.closureTier}${boundaryCls}"${tooltipAttrs}>${volume}</td>`);
    }
    tr.innerHTML = cells.join("");
    fragment.appendChild(tr);
  }

  tableBody.appendChild(fragment);
  const laneCountsArr = getLaneCountsForSelection();
  if (tableWrap) {
    tableWrap.dataset.lanes = laneCountsArr.length === 1 ? String(laneCountsArr[0]) : "";
  }
  const perDateAggregated = isWeekAggregated
    ? (() => {
        const m = new Map();
        for (const row of state.filtered) {
          const key = `${row.dateKey}__${row.time}`;
          const ex = m.get(key);
          if (!ex || row.volume > ex.volume) m.set(key, row);
        }
        return m;
      })()
    : aggregated;
  const datesForChart = isWeekAggregated ? Array.from(new Set(state.filtered.map((r) => r.dateKey))).sort((a, b) => a.localeCompare(b)) : dates;
  state.currentTableData = {
    dates: datesForChart,
    times,
    aggregated: perDateAggregated,
    location: getFilters().location,
    direction: getFilters().direction,
    ...(isWeekAggregated && { weekStarts: dates, weekAggregated: aggregated }),
  };
  tableMeta.textContent = `${formatNumber(
    dates.length
  )} days · ${formatNumber(times.length)} hours · ${formatNumber(
    state.filtered.length
  )} entries`;
  rowCount.textContent = formatNumber(state.filtered.length);
  tableScrollbar.style.display = "block";
  tableScrollbarInner.style.width = `${tableWrap.scrollWidth}px`;

  if (naSummaryCount) {
    naSummaryCount.textContent = `${formatNumber(naItems.length)} N/A cells`;
  }
  if (naSummaryList) {
    if (naItems.length === 0) {
      naSummaryList.textContent = "No N/A values.";
    } else {
      naItems.sort((a, b) => {
        const dateCmp = a.date.localeCompare(b.date);
        if (dateCmp !== 0) return dateCmp;
        return timeSortKey(a.time) - timeSortKey(b.time);
      });
      const list = naItems
        .map(
          (item) =>
            `<li>${formatDateDisplay(item.date)} @ ${item.time}</li>`
        )
        .join("");
      naSummaryList.innerHTML = `<ul>${list}</ul>`;
    }
  }
}

function setActiveTab(tab) {
  const isProjection = tab === "projections";
  if (currentView) currentView.classList.toggle("is-hidden", isProjection);
  if (projectionsView) projectionsView.classList.toggle("is-hidden", !isProjection);
  if (tabCurrent) tabCurrent.classList.toggle("is-active", !isProjection);
  if (tabProjections) tabProjections.classList.toggle("is-active", isProjection);
  if (isProjection) {
    renderProjectionTable();
  }
}

function getAvailableYears(rows) {
  const years = new Set();
  for (const row of rows) {
    if (row.dateObj) years.add(row.dateObj.getFullYear());
  }
  return Array.from(years).sort((a, b) => a - b);
}

function updateProjectionYearOptions() {
  if (!projectionYear) return;
  const baseYear = getBaseYear(state.rows);
  if (!Number.isFinite(baseYear)) {
    projectionYear.innerHTML = "";
    return;
  }
  const options = [];
  for (let y = baseYear + 1; y <= baseYear + 10; y++) {
    options.push(`<option value="${y}">${y}</option>`);
  }
  projectionYear.innerHTML = options.join("");
  if (
    !state.projection.targetYear ||
    state.projection.targetYear <= baseYear ||
    state.projection.targetYear > baseYear + 10
  ) {
    state.projection.targetYear = baseYear + 1;
  }
  projectionYear.value = String(state.projection.targetYear);
}

function renderProjectionRateInputs() {
  if (!projectionLocationRates) return;
  projectionLocationRates.innerHTML = "";
  const filters = getFilters();
  if (state.rows.length === 0) {
    projectionLocationRates.textContent = "Load data to set location rates.";
    return;
  }
  if (!filters.location || !filters.direction) {
    projectionLocationRates.textContent = "Select location and direction to set growth rate.";
    return;
  }
  const location = filters.location;
  const direction = filters.direction;
  const fragment = document.createDocumentFragment();
  const wrapper = document.createElement("div");
  wrapper.className = "projection-rate-row";
  const label = document.createElement("div");
  label.textContent = `${location} (${direction})`;
  const input = document.createElement("input");
  input.type = "text";
  input.inputMode = "numeric";
  input.dataset.location = location;
  const value =
    state.projection.locationRates[location] ??
    state.projection.defaultRate;
  applyDecimalValue(input, value);
  const updateRate = () => {
    const next = readDecimalInput(input);
    if (Number.isFinite(next)) {
      state.projection.locationRates[location] = next;
    }
    scheduleProjectionRender();
  };
  bindDecimalCaretFix(input, updateRate);
  input.addEventListener("input", updateRate);
  const suffix = document.createElement("span");
  suffix.className = "projection-rate-suffix";
  suffix.textContent = "%";
  wrapper.appendChild(label);
  wrapper.appendChild(input);
  wrapper.appendChild(suffix);
  fragment.appendChild(wrapper);
  projectionLocationRates.appendChild(fragment);
}

function getProjectionRateForLocation(location) {
  const value =
    state.projection.locationRates[location] ?? state.projection.defaultRate;
  return Number.isFinite(value) ? value / 100 : 0;
}

function buildHolidayMappings(baseYear, targetYear) {
  const baseList = buildHolidayList(baseYear);
  const targetList = buildHolidayList(targetYear);
  const baseByLabel = new Map();
  for (const item of baseList) {
    baseByLabel.set(item.label, item.dateObj);
  }
  return targetList
    .map((item) => {
      const baseDate = baseByLabel.get(item.label);
      if (!baseDate) return null;
      return {
        label: item.label,
        baseDate,
        targetDate: item.dateObj,
      };
    })
    .filter(Boolean);
}

function renderHolidaySummary(mappings) {
  if (!projectionHolidaySummary) return;
  if (!mappings || mappings.length === 0) {
    projectionHolidaySummary.textContent = "No holiday mappings available.";
    return;
  }
  mappings.sort((a, b) => a.targetDate - b.targetDate);
  const rows = mappings
    .map((item) => {
      const baseKey = formatDateKey(item.baseDate);
      const targetKey = formatDateKey(item.targetDate);
      const weekday = getWeekdayLabel(String(item.baseDate.getDay()));
      return `<tr>
        <td>${item.label}</td>
        <td>${weekday}</td>
        <td>${formatDateDisplay(baseKey)}</td>
        <td>${formatDateDisplay(targetKey)}</td>
      </tr>`;
    })
    .join("");
  projectionHolidaySummary.innerHTML = `<table>
    <thead>
      <tr>
        <th>Holiday</th>
        <th>Source day</th>
        <th>Source date</th>
        <th>Target date</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function buildTargetToSourceMap(baseYear, targetYear) {
  const holidayMappings = buildHolidayMappings(baseYear, targetYear);
  const targetToSource = new Map();
  for (const item of holidayMappings) {
    const targetKey = formatDateKey(item.targetDate);
    const sourceKey = formatDateKey(item.baseDate);
    targetToSource.set(targetKey, sourceKey);
  }
  return targetToSource;
}

function findSourceDateKey(targetDateKey, targetToSourceHoliday, baseYear, baseHolidayDateKeys) {
  const override = targetToSourceHoliday.get(targetDateKey);
  if (override) return override;
  const parts = targetDateKey.split("-");
  if (parts.length !== 3) return null;
  const targetDate = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );
  let sourceDate = mapDateToYearByWeekday(targetDate, baseYear);
  let sourceKey = formatDateKey(sourceDate);

  // Avoid using a base-year holiday when target is not a holiday
  if (baseHolidayDateKeys && baseHolidayDateKeys.has(sourceKey)) {
    const baseYearStart = new Date(baseYear, 0, 1);
    const baseYearEnd = new Date(baseYear, 11, 31);
    for (const delta of [-7, 7, -14, 14, -21, 21]) {
      const alt = new Date(sourceDate);
      alt.setDate(alt.getDate() + delta);
      if (alt >= baseYearStart && alt <= baseYearEnd) {
        const altKey = formatDateKey(alt);
        if (!baseHolidayDateKeys.has(altKey)) {
          return altKey;
        }
      }
    }
    // Fallback: use -7 even if holiday (least bad)
    sourceDate.setDate(sourceDate.getDate() - 7);
    sourceKey = formatDateKey(sourceDate);
  }
  return sourceKey;
}

function buildProjectedRows(rows, targetYear) {
  if (!rows.length) return [];
  const baseYear = getBaseYear(rows);
  if (!Number.isFinite(baseYear)) return [];

  const baseRows = rows.filter(
    (row) => row.dateObj && row.dateObj.getFullYear() === baseYear
  );
  if (!baseRows.length) return [];

  const targetToSourceHoliday = buildTargetToSourceMap(baseYear, targetYear);

  const baseHolidayDateKeys = new Set();
  for (const item of buildHolidayList(baseYear)) {
    baseHolidayDateKeys.add(formatDateKey(item.dateObj));
  }

  const sourceByKey = new Map();
  for (const row of baseRows) {
    const key = `${row.dateKey}__${row.time}`;
    const existing = sourceByKey.get(key);
    if (!existing || row.volume > existing.volume) {
      sourceByKey.set(key, row);
    }
  }

  const targetDates = getFullYearDateKeys(targetYear);
  const times = Array.from(new Set(baseRows.map((r) => r.time)));
  const rate = getProjectionRateForLocation(
    baseRows[0]?.location || "Unknown"
  );
  const yearDelta = targetYear - baseYear;
  const growth = Math.pow(1 + rate, yearDelta);

  const projected = [];
  for (const targetDateKey of targetDates) {
    const sourceDateKey = findSourceDateKey(
      targetDateKey,
      targetToSourceHoliday,
      baseYear,
      baseHolidayDateKeys
    );
    for (const time of times) {
      const sourceKey = sourceDateKey ? `${sourceDateKey}__${time}` : null;
      const sourceRow = sourceKey ? sourceByKey.get(sourceKey) : null;
      const baseVolume = sourceRow ? sourceRow.volume : Number.NaN;
      const projectedVolume = Number.isNaN(baseVolume)
        ? Number.NaN
        : Math.ceil(baseVolume * growth);
      const lanes = sourceRow ? sourceRow.lanes : null;
      projected.push({
        location: sourceRow?.location || baseRows[0]?.location || "Unknown",
        direction:
          sourceRow?.direction || baseRows[0]?.direction || "Unknown",
        dateKey: targetDateKey,
        dateObj: new Date(
          Number(targetDateKey.slice(0, 4)),
          Number(targetDateKey.slice(5, 7)) - 1,
          Number(targetDateKey.slice(8, 10))
        ),
        time,
        volume: projectedVolume,
        lanes,
        closureTier: deriveTier(projectedVolume, lanes),
        sourceDateKey: sourceDateKey || null,
        sourceVolume: sourceRow && Number.isFinite(sourceRow.volume) ? sourceRow.volume : null,
      });
    }
  }
  return projected;
}

function renderProjectionTable() {
  if (!projectionTableBody || !projectionTableHead) return;
  projectionTableBody.innerHTML = "";
  projectionTableHead.innerHTML = "";

  const filters = getFilters();
  if (state.rows.length === 0) {
    if (projectionTableMeta) {
      projectionTableMeta.textContent = "No data loaded.";
    }
    if (projectionTableScrollbar) {
      projectionTableScrollbar.style.display = "none";
    }
    if (projectionTableScrollbarInner) {
      projectionTableScrollbarInner.style.width = "0px";
    }
    if (projectionHolidaySummary) {
      projectionHolidaySummary.textContent = "Load data to see holiday mappings.";
    }
    return;
  }

  if (!filters.location || !filters.direction) {
    if (projectionTableMeta) {
      projectionTableMeta.textContent =
        "Select location and direction to view projections.";
    }
    if (projectionTableScrollbar) {
      projectionTableScrollbar.style.display = "none";
    }
    if (projectionTableScrollbarInner) {
      projectionTableScrollbarInner.style.width = "0px";
    }
    return;
  }

  const targetYearValue = Number(projectionYear?.value);
  if (!Number.isFinite(targetYearValue)) return;
  state.projection.targetYear = targetYearValue;
  const projectionBaseRows = getRowsForProjection(targetYearValue);
  const projectedRows = buildProjectedRows(projectionBaseRows, targetYearValue);

  let dates = getFullYearDateKeys(targetYearValue);
  if (filters.start || filters.end) {
    const yearStart = new Date(targetYearValue, 0, 1);
    const yearEnd = new Date(targetYearValue, 11, 31, 23, 59, 59, 999);
    const overlaps =
      (!filters.end || filters.end >= yearStart) && (!filters.start || filters.start <= yearEnd);
    if (overlaps) {
      const rangeStart = !filters.start ? yearStart : filters.start < yearStart ? yearStart : filters.start;
      const rangeEnd = !filters.end ? yearEnd : filters.end > yearEnd ? yearEnd : new Date(filters.end);
      rangeEnd.setHours(23, 59, 59, 999);
      dates = dates.filter((dateKey) => {
        const [y, m, d] = dateKey.split("-").map(Number);
        const date = new Date(y, m - 1, d);
        return date >= rangeStart && date <= rangeEnd;
      });
    }
  }
  const isMonThu = filters.weekday === "mon-thu";
  const isMonWed = filters.weekday === "mon-wed";
  const isWeekAggregated = isMonThu || isMonWed;
  const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  if (filters.weekday !== "all") {
    dates = dates.filter((dateKey) => {
      const [y, m, d] = dateKey.split("-").map(Number);
      const dayOfWeek = new Date(y, m - 1, d).getDay();
      if (isMonThu) return dayOfWeek >= 1 && dayOfWeek <= 4;
      if (isMonWed) return dayOfWeek >= 1 && dayOfWeek <= 3;
      return String(dayOfWeek) === filters.weekday;
    });
  }
  if (isWeekAggregated) {
    dates = Array.from(new Set(dates.map((d) => getWeekStartSunday(d)))).sort((a, b) => a.localeCompare(b));
  }
  const times = Array.from(new Set(projectedRows.map((row) => row.time))).sort(
    (a, b) => timeSortKey(a) - timeSortKey(b)
  );

  let aggregated;
  let aggregatedMeta;
  if (isWeekAggregated) {
    aggregated = new Map();
    aggregatedMeta = new Map();
    const dayMax = isMonThu ? 4 : 3;
    for (const row of projectedRows) {
      const d = row.dateObj ? row.dateObj.getDay() : -1;
      if (d < 1 || d > dayMax) continue;
      const weekStart = getWeekStartSunday(row.dateKey);
      const key = `${weekStart}__${row.time}`;
      const vol = Number.isFinite(row.volume) ? row.volume : NaN;
      const existing = aggregated.get(key);
      const existingVol = existing && Number.isFinite(existing.volume) ? existing.volume : -Infinity;
      if (!Number.isNaN(vol) && vol > existingVol) {
        aggregated.set(key, row);
        aggregatedMeta.set(key, {
          dateKey: row.dateKey,
          dayName: DAY_NAMES[d],
          volume: row.volume,
        });
      } else if (!existing) {
        aggregated.set(key, row);
        aggregatedMeta.set(key, {
          dateKey: row.dateKey,
          dayName: DAY_NAMES[d],
          volume: row.volume,
        });
      }
    }
  } else {
    aggregated = new Map();
    for (const row of projectedRows) {
      const key = `${row.dateKey}__${row.time}`;
      const existing = aggregated.get(key);
      if (!existing || row.volume > existing.volume) aggregated.set(key, row);
    }
  }

  const projectionYearSet = new Set([targetYearValue]);
  const projectionHolidayMap = buildHolidayMap(
    projectionYearSet,
    getHolidaySelections()
  );
  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthSpans = [];
  for (let i = 0; i < dates.length; ) {
    const [y, m] = dates[i].split("-").map(Number);
    const monthKey = `${y}-${String(m).padStart(2, "0")}`;
    const startIdx = i;
    let count = 0;
    while (i < dates.length && dates[i].slice(0, 7) === monthKey) {
      count++;
      i++;
    }
    const spanDates = dates.slice(startIdx, startIdx + count);
    let dateKeysForClick = spanDates;
    if (isWeekAggregated) {
      dateKeysForClick = [];
      const dayCount = isMonThu ? 4 : 3;
      for (const weekStart of spanDates) {
        const [yw, mw, dw] = weekStart.split("-").map(Number);
        for (let j = 1; j <= dayCount; j++) {
          const d = new Date(yw, mw - 1, dw + j);
          dateKeysForClick.push(formatDateKey(d));
        }
      }
    }
    monthSpans.push({ name: MONTH_NAMES[m - 1], count, dates: dateKeysForClick });
  }
  const monthRow = document.createElement("tr");
  monthRow.innerHTML = `<th></th>${monthSpans.map((s, idx) => {
    const dateKeys = s.dates.length ? ` data-month-dates="${s.dates.join(",")}"` : "";
    const monthAttr = isMonThu ? ` data-month-mon-thu="true"` : isMonWed ? ` data-month-mon-wed="true"` : "";
    return `<th colspan="${s.count}" class="month-label month-label-clickable${idx > 0 ? " month-boundary" : ""}"${dateKeys}${monthAttr}>${s.name}</th>`;
  }).join("")}`;
  projectionTableHead.appendChild(monthRow);

  const headRow = document.createElement("tr");
  const dateHeaders = dates
    .map((date, i) => {
      const holidays = projectionHolidayMap.get(date) || [];
      const title = holidays.length > 0 ? holidays.join(", ") : "";
      const cls = holidays.length > 0 ? "holiday-header" : "";
      const monthBoundary = i > 0 && dates[i - 1].slice(0, 7) !== date.slice(0, 7);
      const boundaryCls = monthBoundary ? " month-boundary" : "";
      const displayDate = isWeekAggregated ? `Week of ${formatDateDisplay(date)}` : formatDateDisplay(date);
      const [yw, mw, dw] = date.split("-").map(Number);
      const weekDates = isWeekAggregated
        ? Array.from({ length: isMonThu ? 4 : 3 }, (_, j) => formatDateKey(new Date(yw, mw - 1, dw + j + 1)))
        : null;
      const weekDatesAttr = weekDates ? ` data-week-dates="${weekDates.join(",")}"` : "";
      return `<th class="${cls}${boundaryCls} date-clickable" data-date="${date}"${weekDatesAttr} title="${title}">${displayDate}</th>`;
    })
    .join("");
  headRow.innerHTML = `<th>Time</th>${dateHeaders}`;
  projectionTableHead.appendChild(headRow);

  const fragment = document.createDocumentFragment();
  for (const time of times) {
    const tr = document.createElement("tr");
    const cells = [`<td>${time}</td>`];
    for (let i = 0; i < dates.length; i++) {
      const date = dates[i];
      const key = `${date}__${time}`;
      const row = aggregated.get(key);
      const meta = isWeekAggregated ? aggregatedMeta.get(key) : null;
      const monthBoundary = i > 0 && dates[i - 1].slice(0, 7) !== date.slice(0, 7);
      const boundaryCls = monthBoundary ? " month-boundary" : "";
      const weekdayRangeAttr = isMonWed ? ` data-weekday-range="mon-wed"` : "";
      const tooltipAttrs = isWeekAggregated && meta
        ? ` data-max-date="${meta.dateKey}" data-max-day="${(meta.dayName || "").replace(/"/g, "&quot;")}" data-max-volume="${Number.isFinite(meta.volume) ? meta.volume : ""}"${weekdayRangeAttr}`
        : "";
      if (!row) {
        cells.push(`<td class="cell-unknown${boundaryCls}"${tooltipAttrs}>—</td>`);
        continue;
      }
      const sourceAttrs =
        !isWeekAggregated && row.sourceDateKey != null
          ? ` data-source-date="${row.sourceDateKey}" data-source-time="${String(row.time).replace(/"/g, "&quot;")}" data-source-volume="${row.sourceVolume != null ? String(row.sourceVolume) : ""}"`
          : "";
      if (Number.isNaN(row.volume)) {
        cells.push(`<td class="cell-na${boundaryCls}"${tooltipAttrs}${sourceAttrs}>N/A</td>`);
        continue;
      }
      const volume = formatNumber(row.volume);
      cells.push(
        `<td class="cell-${row.closureTier}${boundaryCls}"${tooltipAttrs}${sourceAttrs}>${volume}</td>`
      );
    }
    tr.innerHTML = cells.join("");
    fragment.appendChild(tr);
  }

  projectionTableBody.appendChild(fragment);
  const projLaneCounts = Array.from(new Set(projectedRows.map((r) => r.lanes).filter(Number.isFinite)));
  if (projectionTableWrap) {
    projectionTableWrap.dataset.lanes = projLaneCounts.length === 1 ? String(projLaneCounts[0]) : "";
  }
  const perDateAggregated = isWeekAggregated
    ? (() => {
        const m = new Map();
        const dayMax = isMonThu ? 4 : 3;
        for (const row of projectedRows) {
          const d = row.dateObj ? row.dateObj.getDay() : -1;
          if (d < 1 || d > dayMax) continue;
          const key = `${row.dateKey}__${row.time}`;
          const ex = m.get(key);
          if (!ex || row.volume > ex.volume) m.set(key, row);
        }
        return m;
      })()
    : aggregated;
  const datesForChart = isWeekAggregated
    ? Array.from(new Set(projectedRows.filter((r) => r.dateObj && r.dateObj.getDay() >= 1 && r.dateObj.getDay() <= (isMonThu ? 4 : 3)).map((r) => r.dateKey))).sort((a, b) => a.localeCompare(b))
    : dates;
  state.projectionTableData = {
    dates: datesForChart,
    times,
    aggregated: perDateAggregated,
    location: getFilters().location,
    direction: getFilters().direction,
    isProjection: true,
    ...(isWeekAggregated && { weekStarts: dates, weekAggregated: aggregated }),
  };
  if (projectionTableMeta) {
    projectionTableMeta.textContent = `${formatNumber(
      dates.length
    )} days · ${formatNumber(times.length)} hours · ${formatNumber(
      projectedRows.length
    )} entries`;
  }
  if (projectionTableScrollbar) {
    projectionTableScrollbar.style.display = "block";
  }
  if (projectionTableScrollbarInner) {
    projectionTableScrollbarInner.style.width = `${projectionTableWrap.scrollWidth}px`;
  }

  const baseYear = getBaseYear(state.filtered);
  if (Number.isFinite(baseYear)) {
    renderHolidaySummary(buildHolidayMappings(baseYear, targetYearValue));
  }
}

let projectionRenderTimer = null;
function scheduleProjectionRender() {
  if (projectionRenderTimer) {
    clearTimeout(projectionRenderTimer);
  }
  projectionRenderTimer = setTimeout(() => {
    renderProjectionTable();
    renderProjectionClosureSchedule();
    projectionRenderTimer = null;
  }, 150);
}

function timeSortKey(value) {
  if (!value) return Number.MAX_SAFE_INTEGER;
  const match = value.trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (!match) return Number.MAX_SAFE_INTEGER - 1;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const meridiem = match[3] ? match[3].toUpperCase() : "";
  if (meridiem === "PM" && hour < 12) hour += 12;
  if (meridiem === "AM" && hour === 12) hour = 0;
  return hour * 60 + minute;
}

const HOUR_HUES = [0, 30, 60, 90, 120, 150, 180, 210, 240, 270, 300, 330];

function getHourColor(hour) {
  const h = ((hour % 24) + 24) % 24;
  const hue = HOUR_HUES[Math.floor(h / 2)];
  const alpha = (h % 2 === 0) ? 1 : 0.6;
  return `hsla(${hue}, 75%, 55%, ${alpha})`;
}

function hslToHex(hslStr) {
  const m = hslStr.match(/hsl\((\d+),\s*(\d+)%,\s*(\d+)%\)/);
  if (!m) return null;
  const h = Number(m[1]) / 360;
  const s = Number(m[2]) / 100;
  const l = Number(m[3]) / 100;
  let r, g, b;
  if (s === 0) {
    r = g = b = l;
  } else {
    const hue2rgb = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue2rgb(p, q, h + 1 / 3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1 / 3);
  }
  const toHex = (x) => {
    const hex = Math.round(x * 255).toString(16);
    return hex.length === 1 ? "0" + hex : hex;
  };
  return "#" + toHex(r) + toHex(g) + toHex(b);
}

function getWeekStart(dateKey) {
  const parts = String(dateKey).split("-").map(Number);
  if (parts.length !== 3) return dateKey;
  const [year, month, day] = parts;
  const date = new Date(year, month - 1, day);
  const weekday = date.getDay();
  const sundayOffset = weekday;
  date.setDate(date.getDate() - sundayOffset);
  return formatDateKey(date);
}

/** Monday of the week containing dateKey (for Mon-Thu aggregation). */
function getWeekStartSunday(dateKey) {
  const parts = String(dateKey).split("-").map(Number);
  if (parts.length !== 3) return dateKey;
  const [year, month, day] = parts;
  const date = new Date(year, month - 1, day);
  const weekday = date.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  const mondayOffset = weekday === 0 ? -6 : 1 - weekday;
  date.setDate(date.getDate() + mondayOffset);
  return formatDateKey(date);
}

/** Sunday of the week containing dateKey (for Mon-Thu "Week of" display). */
function getWeekStartSunday(dateKey) {
  const parts = String(dateKey).split("-").map(Number);
  if (parts.length !== 3) return dateKey;
  const [year, month, day] = parts;
  const date = new Date(year, month - 1, day);
  const weekday = date.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  date.setDate(date.getDate() - weekday);
  return formatDateKey(date);
}

function getClosureBeginFromSortedRows(sortedRows, threshold) {
  if (!sortedRows || sortedRows.length < 2) return null;
  let lastTime = null;
  for (let i = 1; i < sortedRows.length; i += 1) {
    const prev = sortedRows[i - 1];
    const curr = sortedRows[i];
    const prevAbove = Number.isFinite(prev.volume) && prev.volume > threshold;
    const currAtOrBelow = !Number.isFinite(curr.volume) || curr.volume <= threshold;
    if (prevAbove && currAtOrBelow) lastTime = curr.time;
  }
  return lastTime ?? "ANY";
}

function getClosureEndFromSortedRows(sortedRows, threshold) {
  if (!sortedRows || sortedRows.length < 2) return null;
  for (let i = 1; i < sortedRows.length; i += 1) {
    const prev = sortedRows[i - 1];
    const curr = sortedRows[i];
    const prevAtOrBelow = !Number.isFinite(prev.volume) || prev.volume <= threshold;
    const currAbove = Number.isFinite(curr.volume) && curr.volume > threshold;
    if (prevAtOrBelow && currAbove) return curr.time;
  }
  return "ANY";
}

function getClosureBeginForDay(rows, dateKey, threshold) {
  const dayRows = rows
    .filter((r) => r.dateKey === dateKey)
    .sort((a, b) => timeSortKey(a.time) - timeSortKey(b.time));
  return getClosureBeginFromSortedRows(dayRows, threshold);
}

function getClosureEndForDay(rows, dateKey, threshold) {
  const dayRows = rows
    .filter((r) => r.dateKey === dateKey)
    .sort((a, b) => timeSortKey(a.time) - timeSortKey(b.time));
  return getClosureEndFromSortedRows(dayRows, threshold);
}

/** Build week-aggregated map (key: weekStart__time) using same logic as Lane Closure Table: max volume per slot across Mon-Thu. */
function buildMonThuWeekAggregated(rows) {
  const aggregated = new Map();
  for (const row of rows) {
    let d = -1;
    if (row.dateObj) d = row.dateObj.getDay();
    else if (row.dateKey) {
      const [yr, mo, dy] = row.dateKey.split("-").map(Number);
      if (yr && mo && dy) d = new Date(yr, mo - 1, dy).getDay();
    }
    if (d < 1 || d > 4) continue;
    const weekStart = getWeekStartSunday(row.dateKey);
    const key = `${weekStart}__${row.time}`;
    const vol = Number.isFinite(row.volume) ? row.volume : NaN;
    const existing = aggregated.get(key);
    const existingVol = existing && Number.isFinite(existing.volume) ? existing.volume : -Infinity;
    if (!Number.isNaN(vol) && vol > existingVol) {
      aggregated.set(key, row);
    } else if (!existing) {
      aggregated.set(key, row);
    }
  }
  return aggregated;
}

/** Closure begin/end from week-aggregated max (key: weekStart__time). Returns times + dateKey/volume for the day that had max at each time. */
function getClosureBeginEndFromWeekAggregated(weekAggregated, weekStart, singleThreshold, doubleThreshold, hasDouble) {
  const prefix = `${weekStart}__`;
  const sortedRows = [];
  const rowByTime = new Map();
  for (const [key, row] of weekAggregated) {
    if (!key.startsWith(prefix)) continue;
    const time = key.slice(prefix.length);
    if (!time) continue;
    sortedRows.push({ time, volume: row?.volume });
    rowByTime.set(time, row);
  }
  sortedRows.sort((a, b) => timeSortKey(a.time) - timeSortKey(b.time));
  const result = {};
  if (Number.isFinite(singleThreshold)) {
    result.beginSingle = getClosureBeginFromSortedRows(sortedRows, singleThreshold);
    result.endSingle = getClosureEndFromSortedRows(sortedRows, singleThreshold);
    const beginRow = result.beginSingle && result.beginSingle !== "ANY" ? rowByTime.get(result.beginSingle) : null;
    const endRow = result.endSingle && result.endSingle !== "ANY" ? rowByTime.get(result.endSingle) : null;
    result.beginSingleDateKey = beginRow?.dateKey ?? null;
    result.beginSingleVolume = beginRow != null && Number.isFinite(beginRow.volume) ? beginRow.volume : null;
    result.endSingleDateKey = endRow?.dateKey ?? null;
    result.endSingleVolume = endRow != null && Number.isFinite(endRow.volume) ? endRow.volume : null;
  }
  if (hasDouble && Number.isFinite(doubleThreshold)) {
    result.beginDouble = getClosureBeginFromSortedRows(sortedRows, doubleThreshold);
    result.endDouble = getClosureEndFromSortedRows(sortedRows, doubleThreshold);
    const beginRow = result.beginDouble && result.beginDouble !== "ANY" ? rowByTime.get(result.beginDouble) : null;
    const endRow = result.endDouble && result.endDouble !== "ANY" ? rowByTime.get(result.endDouble) : null;
    result.beginDoubleDateKey = beginRow?.dateKey ?? null;
    result.beginDoubleVolume = beginRow != null && Number.isFinite(beginRow.volume) ? beginRow.volume : null;
    result.endDoubleDateKey = endRow?.dateKey ?? null;
    result.endDoubleVolume = endRow != null && Number.isFinite(endRow.volume) ? endRow.volume : null;
  }
  return result;
}

/** Mon-Thu aggregated: max volume per time slot across Mon-Thu, then closure begin/end. */
function getClosureBeginEndForMonThuWeek(rows, weekStart, singleThreshold, doubleThreshold, hasDouble) {
  const weekAggregated = buildMonThuWeekAggregated(rows);
  return getClosureBeginEndFromWeekAggregated(weekAggregated, weekStart, singleThreshold, doubleThreshold, hasDouble);
}

/** Build week-aggregated map (key: weekStart__time) using same logic as Lane Closure Table: max volume per slot across Mon-Wed. */
function buildMonWedWeekAggregated(rows) {
  const aggregated = new Map();
  for (const row of rows) {
    let d = -1;
    if (row.dateObj) d = row.dateObj.getDay();
    else if (row.dateKey) {
      const [yr, mo, dy] = row.dateKey.split("-").map(Number);
      if (yr && mo && dy) d = new Date(yr, mo - 1, dy).getDay();
    }
    if (d < 1 || d > 3) continue;
    const weekStart = getWeekStartSunday(row.dateKey);
    const key = `${weekStart}__${row.time}`;
    const vol = Number.isFinite(row.volume) ? row.volume : NaN;
    const existing = aggregated.get(key);
    const existingVol = existing && Number.isFinite(existing.volume) ? existing.volume : -Infinity;
    if (!Number.isNaN(vol) && vol > existingVol) {
      aggregated.set(key, row);
    } else if (!existing) {
      aggregated.set(key, row);
    }
  }
  return aggregated;
}

/** Mon-Wed aggregated: max volume per time slot across Mon-Wed, then closure begin/end. */
function getClosureBeginEndForMonWedWeek(rows, weekStart, singleThreshold, doubleThreshold, hasDouble) {
  const weekAggregated = buildMonWedWeekAggregated(rows);
  return getClosureBeginEndFromWeekAggregated(weekAggregated, weekStart, singleThreshold, doubleThreshold, hasDouble);
}

function buildClosureSchedule(rows, options = {}) {
  const { targetYear = null, weekdayFilter = null } = options;
  if (!rows.length) return { weekStartDates: [], data: new Map(), hasDouble: false, laneCount: null };
  const laneCounts = Array.from(new Set(rows.map((r) => r.lanes).filter(Number.isFinite))).sort((a, b) => a - b);
  const laneCount = laneCounts.length ? laneCounts[0] : null;
  const config = laneCount ? state.thresholds[laneCount] : null;
  const hasDouble = laneCount === 3 && config && Number.isFinite(config.doubleMax);
  const singleThreshold = config?.singleMax;
  const doubleThreshold = config?.doubleMax;

  const weekStarts = new Set();
  for (const row of rows) {
    if (row.dateKey && row.dateKey !== "Unknown") {
      weekStarts.add(getWeekStart(row.dateKey));
    }
  }
  const weekStartDates = Array.from(weekStarts).sort((a, b) => a.localeCompare(b));

  const filters = getFilters();
  let effectiveStart = filters.start;
  let effectiveEnd = filters.end;
  let applyDateFilter = !!(filters.start || filters.end);
  if (targetYear != null && applyDateFilter) {
    const yearStart = new Date(targetYear, 0, 1);
    const yearEnd = new Date(targetYear, 11, 31, 23, 59, 59, 999);
    const overlaps =
      (!filters.end || filters.end >= yearStart) && (!filters.start || filters.start <= yearEnd);
    if (overlaps) {
      effectiveStart = !filters.start ? yearStart : filters.start < yearStart ? yearStart : filters.start;
      effectiveEnd = !filters.end ? yearEnd : filters.end > yearEnd ? yearEnd : new Date(filters.end);
      effectiveEnd.setHours(23, 59, 59, 999);
    } else {
      applyDateFilter = false;
    }
  }
  if (applyDateFilter) {
    const filtered = weekStartDates.filter((wk) => {
      const [y, m, d] = wk.split("-").map(Number);
      const weekStart = new Date(y, m - 1, d);
      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);
      if (effectiveStart && weekEnd < effectiveStart) return false;
      if (effectiveEnd) {
        const end = new Date(effectiveEnd);
        end.setHours(23, 59, 59, 999);
        if (weekStart > end) return false;
      }
      return true;
    });
    weekStartDates.length = 0;
    weekStartDates.push(...filtered);
  }

  const data = new Map();
  const volumeByKey = new Map();
  for (const row of rows) {
    if (row.dateKey && row.time != null) {
      const key = `${row.dateKey}__${row.time}`;
      if (Number.isFinite(row.volume)) volumeByKey.set(key, row.volume);
    }
  }
  const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  for (const weekStart of weekStartDates) {
    const [y, m, d] = weekStart.split("-").map(Number);
    const weekData = {};
    if (weekdayFilter === "mon-thu") {
      weekData["Mon-Thu"] = getClosureBeginEndForMonThuWeek(rows, weekStart, singleThreshold, doubleThreshold, hasDouble);
    }
    if (weekdayFilter === "mon-wed") {
      weekData["Mon-Wed"] = getClosureBeginEndForMonWedWeek(rows, weekStart, singleThreshold, doubleThreshold, hasDouble);
    }
    for (let dayOfWeek = 0; dayOfWeek < 7; dayOfWeek += 1) {
      const date = new Date(y, m - 1, d + dayOfWeek);
      const dateKey = formatDateKey(date);
      const dayData = {};
      if (Number.isFinite(singleThreshold)) {
        dayData.beginSingle = getClosureBeginForDay(rows, dateKey, singleThreshold);
        dayData.endSingle = getClosureEndForDay(rows, dateKey, singleThreshold);
      }
      if (hasDouble && Number.isFinite(doubleThreshold)) {
        dayData.beginDouble = getClosureBeginForDay(rows, dateKey, doubleThreshold);
        dayData.endDouble = getClosureEndForDay(rows, dateKey, doubleThreshold);
      }
      weekData[WEEKDAYS[dayOfWeek]] = dayData;
    }
    data.set(weekStart, weekData);
  }

  return { weekStartDates, data, hasDouble, laneCount, singleThreshold, doubleThreshold, volumeByKey };
}

function renderClosureScheduleTable(schedule, beginHead, beginBody, endHead, endBody, type, weekdayFilter) {
  if (!beginHead || !beginBody || !endHead || !endBody) return;
  beginHead.innerHTML = "";
  beginBody.innerHTML = "";
  endHead.innerHTML = "";
  endBody.innerHTML = "";
  const useDouble = type === "double" && schedule.hasDouble;
  const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const MON_THU = ["Monday", "Tuesday", "Wednesday", "Thursday"];
  const MON_WED = ["Monday", "Tuesday", "Wednesday"];
  const wf = weekdayFilter ?? getFilters().weekday;

  let displayCols;
  if (wf === "mon-thu") {
    displayCols = [{ key: "Mon-Thu", days: MON_THU, merge: true }];
  } else if (wf === "mon-wed") {
    displayCols = [{ key: "Mon-Wed", days: MON_WED, merge: true }];
  } else if (wf && wf !== "all") {
    const idx = Number(wf);
    if (idx >= 0 && idx <= 6) displayCols = [{ key: WEEKDAYS[idx], days: [WEEKDAYS[idx]], merge: false }];
    else displayCols = WEEKDAYS.map((d) => ({ key: d, days: [d], merge: false }));
  } else {
    displayCols = WEEKDAYS.map((d) => ({ key: d, days: [d], merge: false }));
  }

  function timeToMinutes(timeStr) {
    if (!timeStr || timeStr === "ANY") return null;
    return timeSortKey(timeStr);
  }

  function getTimeColor(timeStr) {
    const mins = timeToMinutes(timeStr);
    if (mins == null) return null;
    const hour = Math.floor(mins / 60) % 24;
    return getHourColor(hour);
  }

  function getMergedValue(weekData, cellType, weekStart, days) {
    const vals = days.map((d) => {
      const dd = weekData?.[d];
      return dd ? (useDouble ? (cellType === "begin" ? dd.beginDouble : dd.endDouble) : (cellType === "begin" ? dd.beginSingle : dd.endSingle)) : null;
    }).filter((x) => x != null);
    if (vals.length === 0) return null;
    if (vals.every((x) => x === "ANY")) return "ANY";
    const nonAny = vals.filter((x) => x !== "ANY");
    if (nonAny.length === 0) return "ANY";
    if (cellType === "begin") {
      const maxKey = Math.max(...nonAny.map((t) => timeSortKey(t)));
      return nonAny.find((t) => timeSortKey(t) === maxKey) ?? nonAny[0];
    } else {
      const minKey = Math.min(...nonAny.map((t) => timeSortKey(t)));
      return nonAny.find((t) => timeSortKey(t) === minKey) ?? nonAny[0];
    }
  }

  function formatCell(dayData, cellType, dateKey, tooltipOverride) {
    const val = dayData
      ? (useDouble
          ? (cellType === "begin" ? dayData.beginDouble : dayData.endDouble)
          : (cellType === "begin" ? dayData.beginSingle : dayData.endSingle))
      : null;
    const v = val ?? null;
    const label = cellType === "begin" ? "Closure Begin" : "Closure End";
    let volStr = "";
    const tooltipDateKey = tooltipOverride?.dateKey ?? dateKey;
    const tooltipVolume = tooltipOverride?.volume;
    if (tooltipVolume != null && Number.isFinite(tooltipVolume) && v && v !== "ANY") {
      volStr = `Volume at ${v}: ${Number(tooltipVolume).toLocaleString()}`;
    } else if (schedule.volumeByKey && tooltipDateKey) {
      const volAtTime = v && v !== "ANY" ? schedule.volumeByKey.get(`${tooltipDateKey}__${v}`) : null;
      const dayVols = v && v !== "ANY" ? [] : Array.from(schedule.volumeByKey.keys())
        .filter((k) => k.startsWith(tooltipDateKey + "__"))
        .map((k) => schedule.volumeByKey.get(k))
        .filter((x) => Number.isFinite(x));
      volStr = volAtTime != null
        ? `Volume at ${v}: ${Number(volAtTime).toLocaleString()}`
        : dayVols.length > 0
          ? `Volume range: ${Math.min(...dayVols).toLocaleString()} – ${Math.max(...dayVols).toLocaleString()}`
          : "No volume data";
    }
    const attrsDateKey = tooltipDateKey ?? dateKey;
    const dataAttrs = attrsDateKey
      ? ` data-closure-date="${attrsDateKey}" data-closure-label="${label}: ${v ?? "No data"}" data-closure-volume="${(volStr || "—").replace(/"/g, "&quot;")}"`
      : "";
    if (v == null) return `<td class="closure-cell-no-data" bgcolor="#cbd5e1"${dataAttrs}></td>`;
    if (v === "ANY") return `<td class="closure-cell-any" bgcolor="#e2e8f0"${dataAttrs}>ANY</td>`;
    const color = getTimeColor(v);
    const hex = color ? hslToHex(color) : null;
    const style = color ? ` style="background: ${color};"` : "";
    const bgcolor = hex ? ` bgcolor="${hex}"` : "";
    return `<td${style}${bgcolor}${dataAttrs}>${v}</td>`;
  }

  const headRow = (headEl) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<th>Week</th>${displayCols.map((c) => `<th>${c.key}</th>`).join("")}`;
    headEl.appendChild(tr);
  };
  headRow(beginHead);
  headRow(endHead);

  const beginFragment = document.createDocumentFragment();
  const endFragment = document.createDocumentFragment();
  for (const weekStart of schedule.weekStartDates) {
    const [y, m, d] = weekStart.split("-").map(Number);
    const weekData = schedule.data.get(weekStart);
    const beginRow = document.createElement("tr");
    const beginCells = [`<td class="closure-section-header">${formatDateDisplay(weekStart)}</td>`];
    const endRow = document.createElement("tr");
    const endCells = [`<td class="closure-section-header">${formatDateDisplay(weekStart)}</td>`];
    for (const col of displayCols) {
      let cellBegin, cellEnd, dateKeyForTooltip;
      if (col.merge) {
        const mergedDayData = weekData?.[col.key] ?? (() => {
          const vBegin = getMergedValue(weekData, "begin", weekStart, col.days);
          const vEnd = getMergedValue(weekData, "end", weekStart, col.days);
          return { beginSingle: vBegin, endSingle: vEnd, beginDouble: vBegin, endDouble: vEnd };
        })();
        const firstDate = new Date(y, m - 1, d + WEEKDAYS.indexOf(col.days[0]));
        dateKeyForTooltip = formatDateKey(firstDate);
        const beginTooltip = mergedDayData.beginSingleDateKey != null
          ? { dateKey: mergedDayData.beginSingleDateKey, volume: mergedDayData.beginSingleVolume }
          : (useDouble && mergedDayData.beginDoubleDateKey != null
            ? { dateKey: mergedDayData.beginDoubleDateKey, volume: mergedDayData.beginDoubleVolume }
            : null);
        const endTooltip = mergedDayData.endSingleDateKey != null
          ? { dateKey: mergedDayData.endSingleDateKey, volume: mergedDayData.endSingleVolume }
          : (useDouble && mergedDayData.endDoubleDateKey != null
            ? { dateKey: mergedDayData.endDoubleDateKey, volume: mergedDayData.endDoubleVolume }
            : null);
        cellBegin = formatCell(mergedDayData, "begin", dateKeyForTooltip, beginTooltip);
        cellEnd = formatCell(mergedDayData, "end", dateKeyForTooltip, endTooltip);
      } else {
        const dayName = col.days[0];
        const i = WEEKDAYS.indexOf(dayName);
        const date = new Date(y, m - 1, d + i);
        dateKeyForTooltip = formatDateKey(date);
        cellBegin = formatCell(weekData?.[dayName], "begin", dateKeyForTooltip);
        cellEnd = formatCell(weekData?.[dayName], "end", dateKeyForTooltip);
      }
      beginCells.push(cellBegin);
      endCells.push(cellEnd);
    }
    beginRow.innerHTML = beginCells.join("");
    beginFragment.appendChild(beginRow);
    endRow.innerHTML = endCells.join("");
    endFragment.appendChild(endRow);
  }
  beginBody.appendChild(beginFragment);
  endBody.appendChild(endFragment);
}

function updateClosureScheduleToggles(hasDouble) {
  if (!hasDouble) state.closureScheduleType = "single";
  const type = state.closureScheduleType;
  if (closureScheduleToggles) closureScheduleToggles.style.display = "flex";
  if (closureToggleSingle) closureToggleSingle.classList.toggle("is-active", type === "single");
  if (closureToggleDouble) closureToggleDouble.classList.toggle("is-active", type === "double");
  if (closureToggleDouble) closureToggleDouble.style.display = hasDouble ? "" : "none";
  if (projectionClosureToggles) projectionClosureToggles.style.display = "flex";
  if (projectionClosureToggleSingle) projectionClosureToggleSingle.classList.toggle("is-active", type === "single");
  if (projectionClosureToggleDouble) projectionClosureToggleDouble.classList.toggle("is-active", type === "double");
  if (projectionClosureToggleDouble) projectionClosureToggleDouble.style.display = hasDouble ? "" : "none";
}

function formatHourLabel(hour) {
  if (hour === 0) return "12:00 AM";
  if (hour === 12) return "12:00 PM";
  return hour < 12 ? `${hour}:00 AM` : `${hour - 12}:00 PM`;
}

function renderClosureLegend(el) {
  if (!el) return;
  const HOURS = [0, 4, 8, 12, 16, 20];
  const gradientStops = [];
  for (let h = 0; h < 24; h++) {
    const c = getHourColor(h);
    const pct = (h / 24) * 100;
    const pctNext = ((h + 1) / 24) * 100;
    gradientStops.push(`${c} ${pct}%`, `${c} ${pctNext}%`);
  }
  const labels = HOURS.map(
    (h) => `<span class="closure-schedule-colorbar-label" style="left:${(h / 24) * 100}%">${formatHourLabel(h)}</span>`
  ).join("");
  el.innerHTML = `
    <span class="closure-schedule-legend-title">Time</span>
    <div class="closure-schedule-colorbar">
      <div class="closure-schedule-colorbar-strip" style="background:linear-gradient(to right, ${gradientStops.join(", ")})"></div>
      <div class="closure-schedule-colorbar-labels">${labels}</div>
    </div>
    <span class="closure-schedule-colorbar-extra">
      <span><span class="closure-schedule-legend-swatch closure-schedule-legend-any"></span> ANY</span>
      <span><span class="closure-schedule-legend-swatch closure-schedule-legend-nodata"></span> No data</span>
    </span>
  `;
}

function renderClosureSchedule() {
  if (!closureScheduleBeginHead || !closureScheduleBeginBody || !closureScheduleMeta) return;
  closureScheduleBeginHead.innerHTML = "";
  closureScheduleBeginBody.innerHTML = "";
  if (closureScheduleEndHead) closureScheduleEndHead.innerHTML = "";
  if (closureScheduleEndBody) closureScheduleEndBody.innerHTML = "";

  if (state.rows.length === 0) {
    closureScheduleMeta.textContent = "Load data to view closure schedule.";
    if (closureScheduleHint) closureScheduleHint.style.display = "none";
    if (closureScheduleToggles) closureScheduleToggles.style.display = "none";
    if (closureScheduleTablesWrap) closureScheduleTablesWrap.style.display = "none";
    if (closureScheduleLegend) { closureScheduleLegend.innerHTML = ""; closureScheduleLegend.style.display = "none"; }
    return;
  }

  if (state.filtered.length === 0) {
    closureScheduleMeta.textContent = "Select location and direction.";
    if (closureScheduleHint) closureScheduleHint.style.display = "none";
    if (closureScheduleToggles) closureScheduleToggles.style.display = "none";
    if (closureScheduleTablesWrap) closureScheduleTablesWrap.style.display = "none";
    if (closureScheduleLegend) { closureScheduleLegend.innerHTML = ""; closureScheduleLegend.style.display = "none"; }
    return;
  }

  const filters = getFilters();
  const schedule = buildClosureSchedule(state.filtered, { weekdayFilter: filters.weekday });
  if (schedule.weekStartDates.length === 0) {
    closureScheduleMeta.textContent = "No week data in filtered range.";
    if (closureScheduleHint) closureScheduleHint.style.display = "none";
    if (closureScheduleToggles) closureScheduleToggles.style.display = "none";
    if (closureScheduleTablesWrap) closureScheduleTablesWrap.style.display = "none";
    if (closureScheduleLegend) { closureScheduleLegend.innerHTML = ""; closureScheduleLegend.style.display = "none"; }
    return;
  }

  if (closureScheduleHint) closureScheduleHint.style.display = "block";
  updateClosureScheduleToggles(schedule.hasDouble);
  if (closureScheduleTablesWrap) closureScheduleTablesWrap.style.display = "flex";
  if (closureScheduleLegend) closureScheduleLegend.style.display = "flex";

  closureScheduleMeta.textContent = `${schedule.weekStartDates.length} weeks · ${filters.location || "—"} ${filters.direction || "—"}`;

  renderClosureScheduleTable(
    schedule,
    closureScheduleBeginHead,
    closureScheduleBeginBody,
    closureScheduleEndHead,
    closureScheduleEndBody,
    state.closureScheduleType,
    filters.weekday
  );
  renderClosureLegend(closureScheduleLegend);
}

function getProjectedRowsForSchedule() {
  const filters = getFilters();
  if (!filters.location || !filters.direction || !state.rows.length) return null;
  const targetYear = Number(projectionYear?.value);
  if (!Number.isFinite(targetYear)) return null;
  const baseRows = getRowsForProjection(targetYear);
  return buildProjectedRows(baseRows, targetYear);
}

function renderProjectionClosureSchedule() {
  if (!projectionClosureScheduleBeginHead || !projectionClosureScheduleBeginBody || !projectionClosureScheduleMeta) return;
  projectionClosureScheduleBeginHead.innerHTML = "";
  projectionClosureScheduleBeginBody.innerHTML = "";
  if (projectionClosureScheduleEndHead) projectionClosureScheduleEndHead.innerHTML = "";
  if (projectionClosureScheduleEndBody) projectionClosureScheduleEndBody.innerHTML = "";

  if (state.rows.length === 0) {
    projectionClosureScheduleMeta.textContent = "Load data to view projected closure schedule.";
    if (projectionClosureScheduleHint) projectionClosureScheduleHint.style.display = "none";
    if (projectionClosureToggles) projectionClosureToggles.style.display = "none";
    if (projectionClosureScheduleTablesWrap) projectionClosureScheduleTablesWrap.style.display = "none";
    if (projectionClosureScheduleLegend) { projectionClosureScheduleLegend.innerHTML = ""; projectionClosureScheduleLegend.style.display = "none"; }
    return;
  }

  const filters = getFilters();
  if (!filters.location || !filters.direction) {
    projectionClosureScheduleMeta.textContent = "Select location and direction.";
    if (projectionClosureScheduleHint) projectionClosureScheduleHint.style.display = "none";
    if (projectionClosureToggles) projectionClosureToggles.style.display = "none";
    if (projectionClosureScheduleTablesWrap) projectionClosureScheduleTablesWrap.style.display = "none";
    if (projectionClosureScheduleLegend) { projectionClosureScheduleLegend.innerHTML = ""; projectionClosureScheduleLegend.style.display = "none"; }
    return;
  }

  const projectedRows = getProjectedRowsForSchedule();
  if (!projectedRows || !projectedRows.length) {
    projectionClosureScheduleMeta.textContent = "Select target year to view projected closure schedule.";
    if (projectionClosureScheduleHint) projectionClosureScheduleHint.style.display = "none";
    if (projectionClosureToggles) projectionClosureToggles.style.display = "none";
    if (projectionClosureScheduleTablesWrap) projectionClosureScheduleTablesWrap.style.display = "none";
    if (projectionClosureScheduleLegend) { projectionClosureScheduleLegend.innerHTML = ""; projectionClosureScheduleLegend.style.display = "none"; }
    return;
  }

  const targetYear = Number(projectionYear?.value);
  const schedule = buildClosureSchedule(projectedRows, { targetYear, weekdayFilter: filters.weekday });
  if (schedule.weekStartDates.length === 0) {
    projectionClosureScheduleMeta.textContent = "No week data in projected range.";
    if (projectionClosureScheduleHint) projectionClosureScheduleHint.style.display = "none";
    if (projectionClosureToggles) projectionClosureToggles.style.display = "none";
    if (projectionClosureScheduleTablesWrap) projectionClosureScheduleTablesWrap.style.display = "none";
    if (projectionClosureScheduleLegend) { projectionClosureScheduleLegend.innerHTML = ""; projectionClosureScheduleLegend.style.display = "none"; }
    return;
  }

  if (projectionClosureScheduleHint) projectionClosureScheduleHint.style.display = "block";
  updateClosureScheduleToggles(schedule.hasDouble);
  if (projectionClosureScheduleTablesWrap) projectionClosureScheduleTablesWrap.style.display = "flex";
  if (projectionClosureScheduleLegend) projectionClosureScheduleLegend.style.display = "flex";

  projectionClosureScheduleMeta.textContent = `${schedule.weekStartDates.length} weeks · ${filters.location || "—"} ${filters.direction || "—"} · ${targetYear} (projected)`;

  renderClosureScheduleTable(
    schedule,
    projectionClosureScheduleBeginHead,
    projectionClosureScheduleBeginBody,
    projectionClosureScheduleEndHead,
    projectionClosureScheduleEndBody,
    state.closureScheduleType,
    filters.weekday
  );
  renderClosureLegend(projectionClosureScheduleLegend);
}

function getHolidaySelections() {
  return {
    us: holidayUS.checked,
    me: holidayME.checked,
    nh: holidayNH.checked,
    disabledHolidays: state.holidays.disabledHolidays,
  };
}

function nthWeekdayOfMonth(year, monthIndex, weekday, n) {
  const first = new Date(year, monthIndex, 1);
  const firstWeekday = first.getDay();
  const offset = (weekday - firstWeekday + 7) % 7;
  const day = 1 + offset + (n - 1) * 7;
  return new Date(year, monthIndex, day);
}

function lastWeekdayOfMonth(year, monthIndex, weekday) {
  const last = new Date(year, monthIndex + 1, 0);
  const lastWeekday = last.getDay();
  const offset = (lastWeekday - weekday + 7) % 7;
  return new Date(year, monthIndex, last.getDate() - offset);
}

function addHoliday(map, dateObj, label) {
  const key = formatDateKey(dateObj);
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(label);
}

function addObservedFixedHoliday(map, year, monthIndex, day, label) {
  const actual = new Date(year, monthIndex, day);
  addHoliday(map, actual, label);
  const weekday = actual.getDay();
  if (weekday === 6) {
    const observed = new Date(year, monthIndex, day - 1);
    addHoliday(map, observed, `${label} (Observed)`);
  }
  if (weekday === 0) {
    const observed = new Date(year, monthIndex, day + 1);
    addHoliday(map, observed, `${label} (Observed)`);
  }
}

function isHolidayEnabled(label, selections) {
  const disabled = selections.disabledHolidays;
  if (!disabled || !disabled.size) return true;
  if (disabled.has(label)) return false;
  if (label.endsWith(" (Observed)")) {
    return !disabled.has(label.replace(" (Observed)", ""));
  }
  return true;
}

function buildHolidayMap(years, selections) {
  const map = new Map();
  const add = (dateObj, label) => {
    if (!isHolidayEnabled(label, selections)) return;
    addHoliday(map, dateObj, label);
  };
  const addObserved = (year, monthIndex, day, label) => {
    if (!isHolidayEnabled(label, selections)) return;
    addObservedFixedHoliday(map, year, monthIndex, day, label);
  };
  for (const year of years) {
    if (selections.us) {
      addObserved(year, 0, 1, "New Year's Day");
      add(nthWeekdayOfMonth(year, 0, 1, 3), "Martin Luther King Jr. Day");
      add(nthWeekdayOfMonth(year, 1, 1, 3), "Washington's Birthday");
      add(lastWeekdayOfMonth(year, 4, 1), "Memorial Day");
      addObserved(year, 5, 19, "Juneteenth");
      addObserved(year, 6, 4, "Independence Day");
      add(nthWeekdayOfMonth(year, 8, 1, 1), "Labor Day");
      add(nthWeekdayOfMonth(year, 9, 1, 2), "Columbus Day");
      addObserved(year, 10, 11, "Veterans Day");
      add(nthWeekdayOfMonth(year, 10, 4, 4), "Thanksgiving Day");
      addObserved(year, 11, 25, "Christmas Day");
    }
    if (selections.me) {
      add(nthWeekdayOfMonth(year, 3, 1, 3), "Patriots' Day");
    }
    if (selections.nh) {
      add(nthWeekdayOfMonth(year, 0, 1, 3), "Civil Rights Day");
      add(nthWeekdayOfMonth(year, 3, 1, 3), "Patriots' Day");
    }
  }
  return map;
}

function addHolidayToList(list, dateObj, label) {
  list.push({ label, dateObj });
}

function addObservedFixedHolidayToList(list, year, monthIndex, day, label) {
  const actual = new Date(year, monthIndex, day);
  addHolidayToList(list, actual, label);
  const weekday = actual.getDay();
  if (weekday === 6) {
    const observed = new Date(year, monthIndex, day - 1);
    addHolidayToList(list, observed, `${label} (Observed)`);
  }
  if (weekday === 0) {
    const observed = new Date(year, monthIndex, day + 1);
    addHolidayToList(list, observed, `${label} (Observed)`);
  }
}

function buildHolidayList(year) {
  const list = [];
  addObservedFixedHolidayToList(list, year, 0, 1, "New Year's Day");
  addHolidayToList(
    list,
    nthWeekdayOfMonth(year, 0, 1, 3),
    "Martin Luther King Jr. Day"
  );
  addHolidayToList(
    list,
    nthWeekdayOfMonth(year, 1, 1, 3),
    "Washington's Birthday"
  );
  addHolidayToList(list, lastWeekdayOfMonth(year, 4, 1), "Memorial Day");
  addObservedFixedHolidayToList(list, year, 5, 19, "Juneteenth");
  addObservedFixedHolidayToList(list, year, 6, 4, "Independence Day");
  addHolidayToList(list, nthWeekdayOfMonth(year, 8, 1, 1), "Labor Day");
  addHolidayToList(list, nthWeekdayOfMonth(year, 9, 1, 2), "Columbus Day");
  addObservedFixedHolidayToList(list, year, 10, 11, "Veterans Day");
  addHolidayToList(
    list,
    nthWeekdayOfMonth(year, 10, 4, 4),
    "Thanksgiving Day"
  );
  addObservedFixedHolidayToList(list, year, 11, 25, "Christmas Day");
  addHolidayToList(list, nthWeekdayOfMonth(year, 3, 1, 3), "Patriots' Day");
  addHolidayToList(list, nthWeekdayOfMonth(year, 0, 1, 3), "Civil Rights Day");

  const thanksgiving = nthWeekdayOfMonth(year, 10, 4, 4);
  const dayBeforeThanksgiving = new Date(thanksgiving);
  dayBeforeThanksgiving.setDate(thanksgiving.getDate() - 1);
  addHolidayToList(list, dayBeforeThanksgiving, "Day Before Thanksgiving");
  addHolidayToList(list, new Date(year, 11, 24), "Christmas Eve");
  addHolidayToList(list, new Date(year, 11, 31), "New Year's Eve");

  return list;
}

function getWeekdayLabel(value) {
  if (value === "mon-thu") return "Mon-Thu";
  if (value === "mon-wed") return "Mon-Wed";
  if (value === "all") return "All";
  const labels = {
    0: "Sunday",
    1: "Monday",
    2: "Tuesday",
    3: "Wednesday",
    4: "Thursday",
    5: "Friday",
    6: "Saturday",
  };
  return labels[value] || "All";
}

function getLaneCounts(rows) {
  const counts = new Set();
  for (const row of rows) {
    if (Number.isFinite(row.lanes)) counts.add(row.lanes);
  }
  return Array.from(counts).sort((a, b) => a - b);
}

function buildLegendHtml(laneCounts) {
  const counts = laneCounts.length ? laneCounts : [2, 3];
  const parts = [];
  if (counts.includes(2)) {
    const two = state.thresholds[2];
    const t90 = 0.9 * two.singleMax;
    const t95 = 0.95 * two.singleMax;
    parts.push(
      `<span class="legend-item tier-single-ok">2-lane: single closure, &lt; 90% (≤ ${formatNumber(t90)})</span>`
    );
    parts.push(
      `<span class="legend-item tier-single-yellow">2-lane: single closure, 90–95% (${formatNumber(t90)}–${formatNumber(t95)})</span>`
    );
    parts.push(
      `<span class="legend-item tier-single-orange">2-lane: single closure, 95–100% (${formatNumber(t95)}–${formatNumber(two.singleMax)})</span>`
    );
    parts.push(
      `<span class="legend-item tier-none">2-lane: no closure (over ${formatNumber(two.singleMax)})</span>`
    );
  }
  if (counts.includes(3)) {
    const three = state.thresholds[3];
    const whiteAt = 0.9 * (three.doubleMax || 0);
    parts.push(
      `<span class="legend-item tier-double-ok">3-lane: double (≤ 90% = ${formatNumber(Math.round(whiteAt))})</span>`
    );
    parts.push(
      `<span class="legend-item tier-double-yellow">3-lane: double (90–100% = ${formatNumber(Math.round(whiteAt))}–${formatNumber(three.doubleMax)})</span>`
    );
    parts.push(
      `<span class="legend-item tier-single-orange">3-lane: single (${formatNumber(three.doubleMax)}–${formatNumber(three.singleMax)})</span>`
    );
    parts.push(
      `<span class="legend-item tier-none">3-lane: no closure (over ${formatNumber(three.singleMax)})</span>`
    );
  }
  return parts.join("");
}

function isProjectionsViewActive() {
  return projectionsView && !projectionsView.classList.contains("is-hidden");
}

function buildExportHtml() {
  const filters = getFilters();
  const locationLabel = filters.location === "all" ? "All" : filters.location;
  const directionLabel =
    filters.direction === "all" ? "All" : filters.direction;
  const weekdayLabel =
    filters.weekday === "all" ? "All" : getWeekdayLabel(filters.weekday);
  const holidays = [];
  if (holidayUS.checked) holidays.push("US Federal");
  if (holidayME.checked) holidays.push("Maine");
  if (holidayNH.checked) holidays.push("New Hampshire");
  const holidayLabel = holidays.length ? holidays.join(", ") : "None";
  const startLabel = filters.start
    ? filters.start.toISOString().slice(0, 10)
    : "Any";
  const endLabel = filters.end
    ? filters.end.toISOString().slice(0, 10)
    : "Any";

  const isProjection = isProjectionsViewActive();
  const wrap = isProjection ? projectionTableWrap : tableWrap;
  const tableEl = wrap && wrap.querySelector("table");
  const tableHtml = tableEl ? tableEl.outerHTML : "";
  const tableMetaText = isProjection && projectionTableMeta
    ? projectionTableMeta.textContent
    : tableMeta.textContent;
  const closureWrap = isProjection ? projectionClosureScheduleTablesWrap : closureScheduleTablesWrap;
  const closureMetaEl = isProjection ? projectionClosureScheduleMeta : closureScheduleMeta;
  const closureScheduleHtml = closureWrap ? closureWrap.outerHTML : "";
  const closureScheduleMetaText = closureMetaEl ? closureMetaEl.textContent : "";
  const entriesLabel = isProjection ? "Projected" : (rowCount && rowCount.textContent) || "0";
  const viewLabel = isProjection
    ? `Projected (Target year: ${projectionYear ? projectionYear.value : "—"})`
    : "Current";
  const laneCounts = getLaneCounts(
    isProjection ? [] : state.filtered
  );
  const legendHtml = buildLegendHtml(laneCounts);

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Lane Closure Export</title>
    <style>
      body {
        font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
        margin: 24px;
        color: #1b1f23;
      }
      h1 {
        margin: 0 0 6px 0;
        font-size: 20px;
      }
      .meta {
        color: #5b6772;
        font-size: 12px;
        margin-bottom: 16px;
      }
      .filters, .legend {
        margin-bottom: 16px;
      }
      .filters span {
        display: inline-block;
        margin-right: 12px;
        font-size: 12px;
        color: #39414a;
      }
      .legend-item {
        display: inline-block;
        padding: 6px 10px;
        border-radius: 8px;
        font-size: 12px;
        margin-right: 8px;
        border: 1px solid transparent;
      }
      .holiday-header { background: #eef2ff; }
      .holiday-header::after {
        content: "";
        display: block;
        width: 18px;
        height: 3px;
        background: #4f46e5;
        border-radius: 999px;
        margin-top: 4px;
      }
      .tier-double-ok, .tier-single-ok { background: #f8fafc; border-color: #e2e8f0; }
      .tier-double-yellow, .tier-single-yellow { background: #fef08a; border-color: #eab308; }
      .tier-double-orange, .tier-single-orange { background: #fdba74; border-color: #ea580c; }
      .tier-single { background: #fef08a; border-color: #eab308; }
      .tier-none { background: #fca5a5; border-color: #dc2626; }
      table {
        border-collapse: collapse;
        width: 100%;
        font-size: 12px;
      }
      th, td {
        border: 1px solid #e5e7eb;
        padding: 6px 8px;
        text-align: left;
        white-space: nowrap;
      }
      th:first-child, td:first-child { font-weight: 600; }
      .cell-single-ok, .cell-double-ok { background: #f8fafc; }
      .cell-single-yellow, .cell-double-yellow { background: #fef08a; }
      .cell-single-orange, .cell-double-orange { background: #fdba74; }
      .cell-single { background: #fef08a; }
      .cell-none { background: #fca5a5; }
      .cell-unknown { background: #f3f4f6; color: #6b7280; }
      .cell-na { background: #fef2f2; color: #b91c1c; font-weight: 700; }
      .closure-schedule-section { margin-top: 24px; }
      .closure-schedule-section h2 { font-size: 16px; margin: 0 0 12px 0; }
      .closure-schedule-tables { display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 16px; }
      .closure-schedule-wrap { flex: 1; min-width: 200px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; }
      .closure-schedule-table-title { font-weight: 700; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; color: #166534; padding: 8px 10px; background: #dcfce7; border-bottom: 1px solid rgba(22,101,52,0.15); }
      .closure-schedule-wrap table { font-size: 12px; border-collapse: collapse; width: 100%; }
      .closure-schedule-wrap th, .closure-schedule-wrap td { border: 1px solid #e5e7eb; padding: 6px 8px; }
      .closure-schedule-wrap thead th { background: #f0fdf4; }
      .closure-section-header { font-weight: 700; font-size: 11px; text-transform: uppercase; color: #166534; background: #dcfce7 !important; }
      .closure-cell-no-data { background: #cbd5e1; }
      .closure-cell-any { background: #e2e8f0; color: #64748b; }
    </style>
  </head>
  <body>
    <h1>Lane Closure Dashboard Export</h1>
    <div class="meta">Exported: ${new Date().toLocaleString()}</div>
    <div class="filters">
      <span><strong>View:</strong> ${viewLabel}</span>
      <span><strong>Location:</strong> ${locationLabel}</span>
      <span><strong>Direction:</strong> ${directionLabel}</span>
      <span><strong>Day:</strong> ${weekdayLabel}</span>
      <span><strong>Start:</strong> ${startLabel}</span>
      <span><strong>End:</strong> ${endLabel}</span>
      <span><strong>Holidays:</strong> ${holidayLabel}</span>
      <span><strong>Entries:</strong> ${entriesLabel}</span>
    </div>
    <div class="legend">${legendHtml}</div>
    <div class="meta">${tableMetaText}</div>
    ${tableHtml}
    ${closureScheduleHtml ? `
    <div class="closure-schedule-section">
      <h2>Closure Schedule</h2>
      <div class="meta">${closureScheduleMetaText}</div>
      ${closureScheduleHtml}
    </div>
    ` : ""}
  </body>
</html>`;
}

function downloadExport(html, options = {}) {
  const { mime = "text/html;charset=utf-8", extension = "html" } = options;
  const blob = new Blob([html], { type: mime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  link.href = url;
  link.download = `lane-closure-export-${stamp}.${extension}`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function exportToPdf(html) {
  const win = window.open("", "_blank");
  if (!win) {
    alert("Pop-up blocked. Allow pop-ups to export PDF.");
    return;
  }
  win.document.write(html);
  win.document.close();
  win.focus();
  win.print();
}

function handleDataLoad(text) {
  state.rows = parseCSV(text);
  setLocations(state.rows);
  setDirections(state.rows, null);
  updateProjectionYearOptions();
  renderProjectionRateInputs();
  updateRowTiers();
  updateThresholdPanel();
  applyFilters();
  renderTable();
  renderProjectionTable();
  renderClosureSchedule();
  renderProjectionClosureSchedule();
}

const loadingOverlay = document.getElementById("loadingOverlay");

function showLoadingOverlay() {
  if (loadingOverlay) {
    loadingOverlay.classList.add("is-visible");
    loadingOverlay.setAttribute("aria-hidden", "false");
  }
}

function hideLoadingOverlay() {
  if (loadingOverlay) {
    loadingOverlay.classList.remove("is-visible");
    loadingOverlay.setAttribute("aria-hidden", "true");
  }
}

function clearAllData() {
  state.rows = [];
  state.filtered = [];
  if (csvInput) csvInput.value = "";
  if (fileName) fileName.textContent = "No file selected";
  if (tableWrap) tableWrap.dataset.lanes = "";
  if (projectionTableWrap) projectionTableWrap.dataset.lanes = "";
  locationFilter.innerHTML = `<option value="">Select</option>`;
  directionFilter.innerHTML = `<option value="">Select</option>`;
  rowCount.textContent = "0";
  tableHead.innerHTML = "";
  tableBody.innerHTML = "";
  tableMeta.textContent = "No data loaded.";
  tableScrollbar.style.display = "none";
  tableScrollbarInner.style.width = "0px";
  updateThresholdPanel();
  if (projectionTableHead) projectionTableHead.innerHTML = "";
  if (projectionTableBody) projectionTableBody.innerHTML = "";
  if (projectionTableMeta) projectionTableMeta.textContent = "No data loaded.";
  if (projectionTableScrollbar) projectionTableScrollbar.style.display = "none";
  if (projectionTableScrollbarInner) {
    projectionTableScrollbarInner.style.width = "0px";
  }
  if (projectionHolidaySummary) {
    projectionHolidaySummary.textContent = "Load data to see holiday mappings.";
  }
  if (projectionYear) projectionYear.innerHTML = "";
  renderProjectionRateInputs();
  renderClosureSchedule();
  renderProjectionClosureSchedule();
}

let fileLoadId = 0;
let activeFileReader = null;

async function processFile(file) {
  if (!file) return;
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  if (!["csv", "xlsx", "xls"].includes(ext)) {
    alert("Please upload a .csv, .xls, or .xlsx file.");
    return;
  }
  if (activeFileReader) {
    activeFileReader.abort();
    activeFileReader = null;
  }
  fileLoadId += 1;
  const thisLoadId = fileLoadId;
  clearAllData();
  if (fileName) fileName.textContent = file.name;
  showLoadingOverlay();
  try {
    if (ext === "xlsx" || ext === "xls") {
      const ab = await file.arrayBuffer();
      if (thisLoadId !== fileLoadId) return;
      const wb = XLSX.read(ab);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const csv = XLSX.utils.sheet_to_csv(ws);
      handleDataLoad(csv);
    } else {
      await new Promise((resolve, reject) => {
        const reader = new FileReader();
        activeFileReader = reader;
        reader.onload = (e) => {
          activeFileReader = null;
          if (thisLoadId !== fileLoadId) {
            resolve();
            return;
          }
          handleDataLoad(e.target.result);
          resolve();
        };
        reader.onerror = () => {
          activeFileReader = null;
          if (thisLoadId === fileLoadId) reject(reader.error);
          else resolve();
        };
        reader.onabort = () => {
          activeFileReader = null;
          resolve();
        };
        reader.readAsText(file);
      });
    }
  } catch (e) {
    if (thisLoadId !== fileLoadId) return;
    console.error(e);
    alert("Failed to load file. " + (e?.message || ""));
  } finally {
    if (thisLoadId === fileLoadId) hideLoadingOverlay();
  }
}

csvInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    if (fileName) fileName.textContent = "No file selected";
    return;
  }
  await processFile(file);
});

const fileDropZone = document.getElementById("fileDropZone");
if (fileDropZone) {
  fileDropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.stopPropagation();
    fileDropZone.classList.add("is-dragover");
  });
  fileDropZone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (!fileDropZone.contains(e.relatedTarget)) {
      fileDropZone.classList.remove("is-dragover");
    }
  });
  fileDropZone.addEventListener("drop", async (e) => {
    e.preventDefault();
    e.stopPropagation();
    fileDropZone.classList.remove("is-dragover");
    const file = e.dataTransfer?.files?.[0];
    if (file) {
      await processFile(file);
    }
  });
}

clearData.addEventListener("click", () => {
  if (activeFileReader) {
    activeFileReader.abort();
    activeFileReader = null;
  }
  fileLoadId += 1;
  clearAllData();
});

exportView.addEventListener("click", () => {
  const isProjection = isProjectionsViewActive();
  if (state.rows.length === 0) {
    alert("Load data before exporting.");
    return;
  }
  if (!isProjection && state.filtered.length === 0) {
    alert("Select location and direction to export the current view.");
    return;
  }
  if (isProjection) {
    const projTable = projectionTableWrap && projectionTableWrap.querySelector("table");
    if (!projTable || !projTable.querySelector("tbody tr")) {
      alert("Select location and direction to export the projected view.");
      return;
    }
  }
  const html = buildExportHtml();
  const format = exportFormat.value;
  if (format === "pdf") {
    exportToPdf(html);
    return;
  }
  if (format === "excel") {
    downloadExport(html, {
      mime: "application/vnd.ms-excel",
      extension: "xls",
    });
    return;
  }
  downloadExport(html);
});

function handleFilterChange() {
  applyFilters();
  updateThresholdPanel();
  renderProjectionRateInputs();
  renderTable();
  renderProjectionTable();
  renderClosureSchedule();
  renderProjectionClosureSchedule();
}

function onClosureScheduleTypeChange(type) {
  state.closureScheduleType = type;
  if (closureToggleSingle) closureToggleSingle.classList.toggle("is-active", type === "single");
  if (closureToggleDouble) closureToggleDouble.classList.toggle("is-active", type === "double");
  if (projectionClosureToggleSingle) projectionClosureToggleSingle.classList.toggle("is-active", type === "single");
  if (projectionClosureToggleDouble) projectionClosureToggleDouble.classList.toggle("is-active", type === "double");
  renderClosureSchedule();
  renderProjectionClosureSchedule();
}

if (closureToggleSingle) closureToggleSingle.addEventListener("click", () => onClosureScheduleTypeChange("single"));
if (closureToggleDouble) closureToggleDouble.addEventListener("click", () => onClosureScheduleTypeChange("double"));
if (projectionClosureToggleSingle) projectionClosureToggleSingle.addEventListener("click", () => onClosureScheduleTypeChange("single"));
if (projectionClosureToggleDouble) projectionClosureToggleDouble.addEventListener("click", () => onClosureScheduleTypeChange("double"));

locationFilter.addEventListener("change", () => {
  setDirections(state.rows, locationFilter.value || null);
  handleFilterChange();
});
directionFilter.addEventListener("change", () => {
  handleFilterChange();
  setTimeout(handleFilterChange, 100);
});
weekdayFilter.addEventListener("change", () => {
  handleFilterChange();
  setTimeout(handleFilterChange, 100);
});
startDate.addEventListener("change", handleFilterChange);
endDate.addEventListener("change", handleFilterChange);
holidayUS.addEventListener("change", () => {
  syncRegionHolidaysToParent("us");
  renderHolidayDetailList();
  renderTable();
  renderProjectionTable();
  renderProjectionClosureSchedule();
});
holidayME.addEventListener("change", () => {
  syncRegionHolidaysToParent("me");
  renderHolidayDetailList();
  renderTable();
  renderProjectionTable();
  renderProjectionClosureSchedule();
});
holidayNH.addEventListener("change", () => {
  syncRegionHolidaysToParent("nh");
  renderHolidayDetailList();
  renderTable();
  renderProjectionTable();
  renderProjectionClosureSchedule();
});
threshold2Single?.addEventListener("input", refreshThresholdLabelsFromInputs);
threshold3Double?.addEventListener("input", refreshThresholdLabelsFromInputs);
threshold3Single?.addEventListener("input", refreshThresholdLabelsFromInputs);
threshold2Single?.addEventListener("input", scheduleThresholdApply);
threshold3Double?.addEventListener("input", scheduleThresholdApply);
threshold3Single?.addEventListener("input", scheduleThresholdApply);
threshold2Single?.addEventListener("change", updateThresholdsFromInputs);
threshold3Double?.addEventListener("change", updateThresholdsFromInputs);
threshold3Single?.addEventListener("change", updateThresholdsFromInputs);

bindNumericCaretFix(threshold2Single, scheduleThresholdApply);
bindNumericCaretFix(threshold3Double, scheduleThresholdApply);
bindNumericCaretFix(threshold3Single, scheduleThresholdApply);

resetFilters.addEventListener("click", () => {
  locationFilter.value = "";
  directionFilter.value = "";
  weekdayFilter.value = "all";
  startDate.value = "";
  endDate.value = "";
  holidayUS.checked = true;
  holidayME.checked = true;
  holidayNH.checked = true;
  state.holidays.disabledHolidays.clear();
  renderHolidayDetailList();
  handleFilterChange();
});

if (toggleSettings) {
  toggleSettings.addEventListener("click", () => {
    const card = toggleSettings.closest(".card");
    if (!card) return;
    const isCollapsed = card.classList.toggle("collapsed");
    toggleSettings.textContent = isCollapsed ? ">" : "v";
    toggleSettings.setAttribute(
      "aria-label",
      isCollapsed
        ? "Expand closure thresholds"
        : "Collapse closure thresholds"
    );
    toggleSettings.setAttribute("aria-expanded", String(!isCollapsed));
    if (settingsPanel) {
      settingsPanel.setAttribute("aria-hidden", String(isCollapsed));
    }
  });
}

if (toggleHolidays && holidayPanel) {
  toggleHolidays.addEventListener("click", () => {
    const group = toggleHolidays.closest(".checkbox-group");
    if (!group) return;
    const isCollapsed = group.classList.toggle("collapsed");
    toggleHolidays.textContent = isCollapsed ? ">" : "v";
    toggleHolidays.setAttribute(
      "aria-label",
      isCollapsed ? "Expand holidays" : "Collapse holidays"
    );
    toggleHolidays.setAttribute("aria-expanded", String(!isCollapsed));
    holidayPanel.setAttribute("aria-hidden", String(isCollapsed));
  });
}

if (toggleHolidayDetail && holidayDetailList) {
  toggleHolidayDetail.addEventListener("click", () => {
    const isCollapsed = holidayDetailList.classList.toggle("collapsed");
    toggleHolidayDetail.textContent = isCollapsed ? ">" : "v";
    toggleHolidayDetail.setAttribute(
      "aria-label",
      isCollapsed
        ? "Expand to select which holidays to show"
        : "Collapse holiday list"
    );
    toggleHolidayDetail.setAttribute("aria-expanded", String(!isCollapsed));
    holidayDetailList.setAttribute("aria-hidden", String(isCollapsed));
  });
}

if (toggleNaSummary) {
  toggleNaSummary.addEventListener("click", () => {
    const card = toggleNaSummary.closest(".card");
    if (!card) return;
    const isCollapsed = card.classList.toggle("collapsed");
    toggleNaSummary.textContent = isCollapsed ? ">" : "v";
    toggleNaSummary.setAttribute(
      "aria-label",
      isCollapsed ? "Expand N/A summary" : "Collapse N/A summary"
    );
    toggleNaSummary.setAttribute("aria-expanded", String(!isCollapsed));
    if (naSummaryPanel) {
      naSummaryPanel.setAttribute("aria-hidden", String(isCollapsed));
    }
  });
}

if (toggleHolidayMapping && projectionHolidayPanel) {
  toggleHolidayMapping.addEventListener("click", () => {
    const card = toggleHolidayMapping.closest(".card");
    if (!card) return;
    const isCollapsed = card.classList.toggle("collapsed");
    toggleHolidayMapping.textContent = isCollapsed ? ">" : "v";
    toggleHolidayMapping.setAttribute(
      "aria-label",
      isCollapsed
        ? "Expand holiday projection mapping"
        : "Collapse holiday projection mapping"
    );
    toggleHolidayMapping.setAttribute("aria-expanded", String(!isCollapsed));
    projectionHolidayPanel.setAttribute("aria-hidden", String(isCollapsed));
  });
}

if (tabCurrent && tabProjections) {
  tabCurrent.addEventListener("click", () => setActiveTab("current"));
  tabProjections.addEventListener("click", () => setActiveTab("projections"));
}

if (projectionYear) {
  projectionYear.addEventListener("change", () => {
    const value = Number(projectionYear.value);
    if (Number.isFinite(value)) {
      state.projection.targetYear = value;
      renderProjectionTable();
      renderProjectionClosureSchedule();
    }
  });
}

updateThresholdInputs();
updateThresholdPanel();
updateProjectionYearOptions();
renderProjectionRateInputs();
renderProjectionTable();
renderClosureSchedule();
renderProjectionClosureSchedule();
renderHolidayDetailList();
setActiveTab("current");

const projectionSourceTooltip = document.createElement("div");
projectionSourceTooltip.className = "projection-source-tooltip";
projectionSourceTooltip.setAttribute("role", "tooltip");
projectionSourceTooltip.setAttribute("aria-hidden", "true");
document.body.appendChild(projectionSourceTooltip);

const closureScheduleTooltip = document.createElement("div");
closureScheduleTooltip.className = "projection-source-tooltip";
closureScheduleTooltip.setAttribute("role", "tooltip");
closureScheduleTooltip.setAttribute("aria-hidden", "true");
document.body.appendChild(closureScheduleTooltip);

const monThuMaxTooltip = document.createElement("div");
monThuMaxTooltip.className = "projection-source-tooltip";
monThuMaxTooltip.setAttribute("role", "tooltip");
monThuMaxTooltip.setAttribute("aria-hidden", "true");
document.body.appendChild(monThuMaxTooltip);

let projectionSourceTooltipTimer = null;

function showMonThuMaxTooltip(cell) {
  const dateKey = cell.getAttribute("data-max-date");
  if (!dateKey) return;
  const dayName = cell.getAttribute("data-max-day") || "";
  const volStr = cell.getAttribute("data-max-volume");
  const weekdayRange = cell.getAttribute("data-weekday-range");
  const displayDate = formatDateDisplay(dateKey);
  const displayVolume = volStr !== "" && volStr != null && Number.isFinite(Number(volStr)) ? formatNumber(Number(volStr)) : "N/A";
  monThuMaxTooltip.innerHTML = "";
  const title = document.createElement("div");
  title.className = "tooltip-title";
  title.textContent = weekdayRange === "mon-wed" ? "Mon-Wed" : "Mon-Thu";
  monThuMaxTooltip.appendChild(title);
  const row1 = document.createElement("div");
  row1.className = "tooltip-row";
  row1.textContent = `Date: ${displayDate}`;
  monThuMaxTooltip.appendChild(row1);
  const row2 = document.createElement("div");
  row2.className = "tooltip-row";
  row2.textContent = `Day: ${dayName}`;
  monThuMaxTooltip.appendChild(row2);
  const row3 = document.createElement("div");
  row3.className = "tooltip-row";
  row3.textContent = `Volume: ${displayVolume}`;
  monThuMaxTooltip.appendChild(row3);
  const rect = cell.getBoundingClientRect();
  monThuMaxTooltip.style.left = `${rect.left + rect.width / 2}px`;
  monThuMaxTooltip.style.top = `${rect.top - 8}px`;
  monThuMaxTooltip.style.transform = "translate(-50%, -100%)";
  monThuMaxTooltip.classList.add("is-visible");
  monThuMaxTooltip.setAttribute("aria-hidden", "false");
}

function hideMonThuMaxTooltip() {
  monThuMaxTooltip.classList.remove("is-visible");
  monThuMaxTooltip.setAttribute("aria-hidden", "true");
}

function showProjectionSourceTooltip(cell) {
  const dateKey = cell.getAttribute("data-source-date");
  const time = cell.getAttribute("data-source-time");
  const volume = cell.getAttribute("data-source-volume");
  if (!dateKey) return;
  const displayDate = formatDateDisplay(dateKey);
  const displayVolume =
    volume !== "" && volume !== null
      ? formatNumber(Number(volume))
      : "N/A";
  projectionSourceTooltip.innerHTML = "";
  const title = document.createElement("div");
  title.className = "tooltip-title";
  title.textContent = "Source (base year)";
  projectionSourceTooltip.appendChild(title);
  const row1 = document.createElement("div");
  row1.className = "tooltip-row";
  row1.textContent = `Date: ${displayDate}`;
  projectionSourceTooltip.appendChild(row1);
  const row2 = document.createElement("div");
  row2.className = "tooltip-row";
  row2.textContent = `Time: ${time || "—"}`;
  projectionSourceTooltip.appendChild(row2);
  const row3 = document.createElement("div");
  row3.className = "tooltip-row";
  row3.textContent = `Volume: ${displayVolume}`;
  projectionSourceTooltip.appendChild(row3);
  const rect = cell.getBoundingClientRect();
  projectionSourceTooltip.style.left = `${rect.left + rect.width / 2}px`;
  projectionSourceTooltip.style.top = `${rect.top - 8}px`;
  projectionSourceTooltip.style.transform = "translate(-50%, -100%)";
  projectionSourceTooltip.classList.add("is-visible");
  projectionSourceTooltip.setAttribute("aria-hidden", "false");
}

function hideProjectionSourceTooltip() {
  projectionSourceTooltip.classList.remove("is-visible");
  projectionSourceTooltip.setAttribute("aria-hidden", "true");
}

function showClosureScheduleTooltip(cell) {
  const dateKey = cell.getAttribute("data-closure-date");
  if (!dateKey) return;
  const label = cell.getAttribute("data-closure-label");
  const volume = cell.getAttribute("data-closure-volume");
  const displayDate = formatDateDisplay(dateKey);
  closureScheduleTooltip.innerHTML = "";
  const title = document.createElement("div");
  title.className = "tooltip-title";
  title.textContent = "Closure Schedule";
  closureScheduleTooltip.appendChild(title);
  const row1 = document.createElement("div");
  row1.className = "tooltip-row";
  row1.textContent = `Date: ${displayDate}`;
  closureScheduleTooltip.appendChild(row1);
  const row2 = document.createElement("div");
  row2.className = "tooltip-row";
  row2.textContent = label || "";
  closureScheduleTooltip.appendChild(row2);
  if (volume && volume !== "—") {
    const row3 = document.createElement("div");
    row3.className = "tooltip-row";
    row3.textContent = volume;
    closureScheduleTooltip.appendChild(row3);
  }
  const rect = cell.getBoundingClientRect();
  closureScheduleTooltip.style.left = `${rect.left + rect.width / 2}px`;
  closureScheduleTooltip.style.top = `${rect.top - 8}px`;
  closureScheduleTooltip.style.transform = "translate(-50%, -100%)";
  closureScheduleTooltip.classList.add("is-visible");
  closureScheduleTooltip.setAttribute("aria-hidden", "false");
}

function hideClosureScheduleTooltip() {
  closureScheduleTooltip.classList.remove("is-visible");
  closureScheduleTooltip.setAttribute("aria-hidden", "true");
}

function openMonthMonThuDetail(monthDates, tableData, monthName) {
  if (!tableData || !monthDates || monthDates.length === 0) return;
  const volumeByKey = {};
  if (tableData.aggregated) {
    for (const [k, row] of tableData.aggregated) {
      volumeByKey[k] = row && Number.isFinite(row.volume) ? row.volume : null;
    }
  }
  const times = tableData.times || [];
  const weekGroups = {};
  for (const d of monthDates) {
    const w = getWeekStartSunday(d);
    if (!weekGroups[w]) weekGroups[w] = new Set();
    weekGroups[w].add(d);
  }
  const weekStarts = Object.keys(weekGroups).sort((a, b) => a.localeCompare(b));
  const CHART_COLORS = [
    { border: "#15803d", fill: "rgba(22, 128, 61, 0.1)" },
    { border: "#2563eb", fill: "rgba(37, 99, 235, 0.08)" },
    { border: "#dc2626", fill: "rgba(220, 38, 38, 0.08)" },
    { border: "#ca8a04", fill: "rgba(202, 138, 4, 0.08)" },
    { border: "#9333ea", fill: "rgba(147, 51, 234, 0.08)" },
    { border: "#0891b2", fill: "rgba(8, 145, 178, 0.08)" },
  ];
  const weekMaxDatasets = weekStarts.map((weekStart, i) => {
    const allDates = Array.from(weekGroups[weekStart]);
    const data = times.map((time) => {
      const vols = allDates.map((d) => volumeByKey[`${d}__${time}`]);
      const finite = vols.filter((v) => v != null && Number.isFinite(v));
      return finite.length > 0 ? Math.max(...finite) : null;
    });
    const colors = CHART_COLORS[i % CHART_COLORS.length];
    return {
      label: `Week of ${formatDateDisplay(weekStart)}`,
      data,
      borderColor: colors.border,
      backgroundColor: colors.fill,
    };
  });
  const allWeekStarts = tableData.weekStarts || tableData.dates || [];
  const availableWeeks = allWeekStarts.filter((w) => !weekStarts.includes(w)).sort((a, b) => a.localeCompare(b));
  const payload = {
    weekMaxDatasets,
    times,
    volumeByKey,
    availableWeeks,
    location: tableData.location,
    direction: tableData.direction,
    isProjection: !!tableData.isProjection,
    isMonthMonThuMax: true,
    weekdayRange: "mon-thu",
    monthName: monthName || "",
  };
  try {
    localStorage.setItem("laneClosureDateDetail", JSON.stringify(payload));
    window.open("date-detail.html", "_blank");
  } catch (e) {
    console.error(e);
  }
}

function openMonthMonWedDetail(monthDates, tableData, monthName) {
  if (!tableData || !monthDates || monthDates.length === 0) return;
  const volumeByKey = {};
  if (tableData.aggregated) {
    for (const [k, row] of tableData.aggregated) {
      volumeByKey[k] = row && Number.isFinite(row.volume) ? row.volume : null;
    }
  }
  const times = tableData.times || [];
  const weekGroups = {};
  for (const d of monthDates) {
    const w = getWeekStartSunday(d);
    if (!weekGroups[w]) weekGroups[w] = new Set();
    weekGroups[w].add(d);
  }
  const weekStarts = Object.keys(weekGroups).sort((a, b) => a.localeCompare(b));
  const CHART_COLORS = [
    { border: "#15803d", fill: "rgba(22, 128, 61, 0.1)" },
    { border: "#2563eb", fill: "rgba(37, 99, 235, 0.08)" },
    { border: "#dc2626", fill: "rgba(220, 38, 38, 0.08)" },
    { border: "#ca8a04", fill: "rgba(202, 138, 4, 0.08)" },
    { border: "#9333ea", fill: "rgba(147, 51, 234, 0.08)" },
    { border: "#0891b2", fill: "rgba(8, 145, 178, 0.08)" },
  ];
  const weekMaxDatasets = weekStarts.map((weekStart, i) => {
    const allDates = Array.from(weekGroups[weekStart]);
    const data = times.map((time) => {
      const vols = allDates.map((d) => volumeByKey[`${d}__${time}`]);
      const finite = vols.filter((v) => v != null && Number.isFinite(v));
      return finite.length > 0 ? Math.max(...finite) : null;
    });
    const colors = CHART_COLORS[i % CHART_COLORS.length];
    return {
      label: `Week of ${formatDateDisplay(weekStart)}`,
      data,
      borderColor: colors.border,
      backgroundColor: colors.fill,
    };
  });
  const allWeekStarts = tableData.weekStarts || tableData.dates || [];
  const availableWeeks = allWeekStarts.filter((w) => !weekStarts.includes(w)).sort((a, b) => a.localeCompare(b));
  const payload = {
    weekMaxDatasets,
    times,
    volumeByKey,
    availableWeeks,
    location: tableData.location,
    direction: tableData.direction,
    isProjection: !!tableData.isProjection,
    isMonthMonWedMax: true,
    weekdayRange: "mon-wed",
    monthName: monthName || "",
  };
  try {
    localStorage.setItem("laneClosureDateDetail", JSON.stringify(payload));
    window.open("date-detail.html", "_blank");
  } catch (e) {
    console.error(e);
  }
}

function openDateDetail(dateKey, tableData, compareDateKeys, isMonThuMax, isMonWedMax) {
  if (!tableData || !dateKey) return;
  const { dates, times, aggregated, location, direction, isProjection } = tableData;
  const volumeByKey = {};
  if (aggregated) {
    for (const [k, row] of aggregated) {
      volumeByKey[k] = row && Number.isFinite(row.volume) ? row.volume : null;
    }
  }
  const payload = { date: dateKey, dates, times, volumeByKey, location, direction, isProjection: !!isProjection };
  if ((isMonThuMax && compareDateKeys && compareDateKeys.length === 3) || (isMonWedMax && compareDateKeys && compareDateKeys.length === 2)) {
    const allDates = [dateKey, ...compareDateKeys];
    const maxTimeseries = times.map((time) => {
      const vols = allDates.map((d) => volumeByKey[`${d}__${time}`]);
      const finite = vols.filter((v) => v != null && Number.isFinite(v));
      const maxVol = finite.length > 0 ? Math.max(...finite) : null;
      return { time, volume: maxVol };
    });
    payload.timeseries = maxTimeseries;
    payload.isMonThuMax = !!isMonThuMax;
    payload.isMonWedMax = !!isMonWedMax;
    payload.weekdayRange = isMonWedMax ? "mon-wed" : "mon-thu";
    const allWeekStarts = tableData.weekStarts || tableData.dates || [];
    payload.availableWeeks = allWeekStarts.filter((w) => w !== getWeekStartSunday(dateKey)).sort((a, b) => a.localeCompare(b));
    payload.weekStartSunday = getWeekStartSunday(dateKey);
  } else if (compareDateKeys && compareDateKeys.length > 0) {
    payload.compareDateKeys = compareDateKeys;
  }
  try {
    localStorage.setItem("laneClosureDateDetail", JSON.stringify(payload));
    window.open("date-detail.html", "_blank");
  } catch (e) {
    console.error(e);
  }
}

document.addEventListener("click", (e) => {
  const monthCell = e.target.closest("th.month-label-clickable[data-month-dates]");
  if (monthCell) {
    const dateKeysStr = monthCell.getAttribute("data-month-dates");
    if (!dateKeysStr) return;
    const monthDates = dateKeysStr.split(",").filter(Boolean);
    if (monthDates.length === 0) return;
    const tableData = monthCell.closest("#projectionTableWrap") ? state.projectionTableData : state.currentTableData;
    if (!tableData) return;
    if (monthCell.getAttribute("data-month-mon-thu") === "true") {
      const monthName = monthCell.textContent?.trim() || "";
      openMonthMonThuDetail(monthDates, tableData, monthName);
      return;
    }
    if (monthCell.getAttribute("data-month-mon-wed") === "true") {
      const monthName = monthCell.textContent?.trim() || "";
      openMonthMonWedDetail(monthDates, tableData, monthName);
      return;
    }
    const primary = monthDates[0];
    const compare = monthDates.slice(1);
    openDateDetail(primary, tableData, compare);
    return;
  }

  const cell = e.target.closest("th.date-clickable[data-date]");
  if (!cell) return;
  let dateKey = cell.getAttribute("data-date");
  if (!dateKey) return;
  const tableData = cell.closest("#projectionTableWrap") ? state.projectionTableData : state.currentTableData;
  if (!tableData) return;
  const weekDatesStr = cell.getAttribute("data-week-dates");
  let compareDateKeys;
  let isMonThuMax = false;
  let isMonWedMax = false;
  if (weekDatesStr) {
    const weekDates = weekDatesStr.split(",").filter(Boolean);
    dateKey = weekDates[0] || dateKey;
    compareDateKeys = weekDates.slice(1);
    isMonThuMax = compareDateKeys.length === 3;
    isMonWedMax = compareDateKeys.length === 2;
  }
  openDateDetail(dateKey, tableData, compareDateKeys, isMonThuMax, isMonWedMax);
});

let monThuMaxTooltipTimer = null;
if (tableWrap) {
  tableWrap.addEventListener("mouseover", (e) => {
    const cell = e.target.closest("td[data-max-date]");
    if (!cell) return;
    if (monThuMaxTooltipTimer) clearTimeout(monThuMaxTooltipTimer);
    monThuMaxTooltipTimer = setTimeout(() => {
      showMonThuMaxTooltip(cell);
      monThuMaxTooltipTimer = null;
    }, 2000);
  });
  tableWrap.addEventListener("mouseout", (e) => {
    const cell = e.target.closest("td[data-max-date]");
    if (monThuMaxTooltipTimer) {
      clearTimeout(monThuMaxTooltipTimer);
      monThuMaxTooltipTimer = null;
    }
    if (cell && (!e.relatedTarget || !cell.contains(e.relatedTarget))) hideMonThuMaxTooltip();
  });
}

if (projectionTableBody) {
  projectionTableBody.addEventListener("mouseover", (e) => {
    const cell = e.target.closest("td");
    if (!cell) return;
    if (cell.hasAttribute("data-max-date")) {
      if (projectionSourceTooltipTimer) {
        clearTimeout(projectionSourceTooltipTimer);
        projectionSourceTooltipTimer = null;
      }
      hideProjectionSourceTooltip();
      if (monThuMaxTooltipTimer) clearTimeout(monThuMaxTooltipTimer);
      monThuMaxTooltipTimer = setTimeout(() => {
        showMonThuMaxTooltip(cell);
        monThuMaxTooltipTimer = null;
      }, 2000);
      return;
    }
    const sourceCell = e.target.closest("td[data-source-date]");
    if (!sourceCell) return;
    if (projectionSourceTooltipTimer) clearTimeout(projectionSourceTooltipTimer);
    projectionSourceTooltipTimer = setTimeout(() => {
      showProjectionSourceTooltip(sourceCell);
      projectionSourceTooltipTimer = null;
    }, 2000);
  });
  projectionTableBody.addEventListener("mouseout", (e) => {
    if (projectionSourceTooltipTimer) {
      clearTimeout(projectionSourceTooltipTimer);
      projectionSourceTooltipTimer = null;
    }
    if (monThuMaxTooltipTimer) {
      clearTimeout(monThuMaxTooltipTimer);
      monThuMaxTooltipTimer = null;
    }
    const maxCell = e.target.closest("td[data-max-date]");
    if (maxCell && (!e.relatedTarget || !maxCell.contains(e.relatedTarget))) {
      hideMonThuMaxTooltip();
    }
    if (!e.relatedTarget || !projectionTableBody.contains(e.relatedTarget)) {
      hideProjectionSourceTooltip();
    }
  });
}

let closureScheduleTooltipTimer = null;
function setupClosureScheduleTooltip(wrap) {
  if (!wrap) return;
  wrap.addEventListener("mouseover", (e) => {
    const cell = e.target.closest("td[data-closure-date]");
    if (!cell) return;
    if (closureScheduleTooltipTimer) clearTimeout(closureScheduleTooltipTimer);
    closureScheduleTooltipTimer = setTimeout(() => {
      showClosureScheduleTooltip(cell);
      closureScheduleTooltipTimer = null;
    }, 2000);
  });
  wrap.addEventListener("mouseout", (e) => {
    if (closureScheduleTooltipTimer) {
      clearTimeout(closureScheduleTooltipTimer);
      closureScheduleTooltipTimer = null;
    }
    if (!e.relatedTarget || !wrap.contains(e.relatedTarget)) {
      hideClosureScheduleTooltip();
    }
  });
}
setupClosureScheduleTooltip(closureScheduleTablesWrap);
setupClosureScheduleTooltip(projectionClosureScheduleTablesWrap);
