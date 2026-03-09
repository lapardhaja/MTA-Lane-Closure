# Agent Documentation – Lane Closure Dashboard

This document describes implemented behaviors and design decisions so agents (and developers) can maintain consistency and avoid regressions.

---

## 1. Project Purpose

Maine Turnpike Authority lane closure dashboard. Static web app: upload traffic volume data (CSV/Excel), filter by location/direction/date, view volume heatmaps, closure schedules, and year-over-year projections.

---

## 2. Projection Logic (Critical – Do Not Deviate)

### 2.1 Overview

Projections map a **target year** (e.g. 2026) to a **base year** (e.g. 2025) using weekday alignment and holiday rules. Each target date gets a source date; volume is looked up and scaled by growth rate.

### 2.2 Source Date Resolution (`findSourceDateKey`)

Order of operations:

1. **Holiday override** – If the target date is a holiday, use the base-year equivalent holiday (e.g. Labor Day 2026 → Labor Day 2025). `targetToSourceHoliday` map.
2. **Weekday mapping** – Otherwise, use `mapDateToYearByWeekday` to find the nearest same weekday in the base year.
3. **Holiday avoidance** – If the weekday-mapped source date is a base-year holiday, **do not use it**. The target is a non-holiday; using holiday traffic would be wrong.

### 2.3 Holiday Cross-Mapping Avoidance (‑7 / +7 Logic)

**Problem:** A non-holiday target can map to a base-year holiday by weekday. Example: 8/31/2026 (Monday) → 9/1/2025 (Monday, Labor Day). Regular Mondays must not use Labor Day traffic.

**Solution:** When the weekday-mapped source is in `baseHolidayDateKeys`:

1. Try alternatives in order: **‑7, +7, ‑14, +14, ‑21, +21** days.
2. Use the first date that is (a) in the base year, and (b) not a holiday.
3. If none qualify, fall back to **‑7 days** (least bad).

**Examples:**

| Target Date   | Would Map To (Weekday) | Actual Source (After Avoidance) |
|---------------|------------------------|----------------------------------|
| 8/31/2026 Mon | 9/1/2025 (Labor Day)   | 8/25/2025 (prev Monday)         |
| 12/23/2026 Wed| 12/24/2025 (Christmas Eve) | 12/17/2025 (prev Wed)       |
| 9/7/2026 (Labor Day) | —              | 9/1/2025 (holiday override, unchanged) |

**Implementation:** `app.js` – `findSourceDateKey`, `buildProjectedRows`. `baseHolidayDateKeys` is built from `buildHolidayList(baseYear)`.

---

## 3. Holiday Definitions

- **US federal:** New Year's, MLK Jr., Washington's Birthday, Memorial Day, Juneteenth, Independence Day, Labor Day, Columbus Day, Veterans Day, Thanksgiving, Christmas.
- **Maine:** Patriots' Day (configurable).
- **New Hampshire:** Civil Rights Day (configurable).
- **Additional:** Day Before Thanksgiving, Christmas Eve, New Year's Eve.
- **Observed dates:** Fixed holidays that fall on Sat/Sun get observed dates (e.g. New Year's observed).

`buildHolidayList(year)` returns `[{ label, dateObj }]`. All holidays for a year are in this list.

---

## 4. Week Structure

- **Closure schedule / Mon-Thu / Mon-Wed:** Week starts **Sunday** (`getWeekStartSunday`).
- **Mon-Thu aggregation:** Monday–Thursday only; max volume per time slot across those four days.
- **Mon-Wed aggregation:** Monday–Wednesday only; max volume per time slot across those three days.

---

## 5. Closure Schedule Logic

- **Single day:** Closure begin/end from that day’s volume curve.
- **Mon-Thu:** Closure begin/end from the **aggregated max-volume curve** (same as Lane Closure Table), not from merging per-day begin/end. `buildMonThuWeekAggregated` + `getClosureBeginEndForMonThuWeek`.
- **Mon-Wed:** Same logic as Mon-Thu but for Monday–Wednesday only. `buildMonWedWeekAggregated` + `getClosureBeginEndForMonWedWeek`.

---

## 6. Date Filtering for Projections

- **Base-year filter (e.g. 2024):** Source rows filtered by date; projection built from that subset; closure schedule shows all projected weeks.
- **Target-year filter (e.g. 2026):** All source rows used; closure schedule and projection table filtered by the 2026 date range.
- **Start only, no end:** e.g. start 3/1/2026, target 2027 → all source rows used; table shows full 2027.

---

## 7. Key Files

| File            | Role                                      |
|-----------------|-------------------------------------------|
| `app.js`        | All logic: parsing, filters, tables, projections |
| `index.html`    | Main UI, filters, Lane Closure Table, Closure Schedule |
| `date-detail.html` | Volume-by-hour chart, compare dates/weeks |
| `styles.css`    | Styles, theme green `#166534`              |

---

## 8. Invariants (Do Not Break)

1. **Holiday-to-holiday mapping** – Target holiday always maps to base-year same holiday.
2. **Non-holiday never uses holiday source** – Enforced by ‑7/+7 avoidance in `findSourceDateKey`.
3. **Mon-Thu / Mon-Wed closure = aggregated max curve** – Not a merge of per-day begin/end.
4. **Week = Sunday start** – For Mon-Thu, Mon-Wed, and closure schedule.
5. **Base year inferred from data** – `getBaseYear(rows)`; no hardcoded year.

---

## 9. Adding or Changing Holidays

- Edit `buildHolidayList` in `app.js`.
- `HOLIDAYS_BY_REGION` controls which holidays are included per region (US, ME, NH).
- Maine Statehood and Maine Observed were removed per project requirements.
