"use strict";

var dayKeys = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"];

function parseWorkbook(workbook, fileName, fileOrder) {
  var records = [];
  var weekIndex = 0;
  var order = typeof fileOrder === "number" ? fileOrder : 0;

  workbook.SheetNames.forEach(function (sheetName, sheetOrder) {
    var sheet = workbook.Sheets[sheetName];
    var rows = workbookToRows(workbook, sheet);
    var headers = detectHeaderRows(rows);

    headers.forEach(function (header, idx) {
      weekIndex += 1;
      var nextHeaderRow = idx < headers.length - 1 ? headers[idx + 1].row : rows.length;
      var weekLabel = findWeekLabel(rows, header.row, sheetName);
      var dateContextLabel = findDateContextLabel(rows, header.row, weekLabel);
      var weekMeta = deriveWeekMeta(weekLabel);
      records = records.concat(
        extractAssignments(
          rows,
          header,
          nextHeaderRow,
          fileName,
          order,
          sheetName,
          sheetOrder,
          weekLabel,
          dateContextLabel,
          weekIndex,
          weekMeta
        )
      );
    });
  });

  return records;
}

function workbookToRows(workbook, sheet) {
  var XLSX = getXlsx(workbook);
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
}

function getXlsx(workbook) {
  if (workbook && workbook.utils && workbook.utils.sheet_to_json) return workbook;
  if (typeof require === "function") return require("xlsx");
  if (typeof window !== "undefined" && window.XLSX) return window.XLSX;
  throw new Error("XLSX library unavailable");
}

function detectHeaderRows(rows) {
  var headers = [];

  for (var r = 0; r < rows.length; r += 1) {
    var row = rows[r] || [];
    var dayColumns = [];
    var observationCol = -1;

    for (var c = 0; c < row.length; c += 1) {
      var value = toText(row[c]);
      if (!value) continue;
      var n = normalize(value);
      if (observationCol < 0 && (n === "OBSERVATIONS" || n.indexOf("OBSERVATION") === 0)) {
        observationCol = c;
      }
      for (var d = 0; d < dayKeys.length; d += 1) {
        if (n.indexOf(dayKeys[d]) === 0) {
          dayColumns.push({ col: c, dayName: dayKeys[d], dayLabel: value.trim(), dayOrder: d });
          break;
        }
      }
    }

    if (dayColumns.length >= 5) {
      dayColumns.sort(function (a, b) {
        return a.col - b.col;
      });
      headers.push({ row: r, dayColumns: dayColumns, observationCol: observationCol });
    }
  }

  return headers;
}

function findWeekLabel(rows, headerRow, fallback) {
  for (var i = headerRow; i >= Math.max(0, headerRow - 6); i -= 1) {
    var row = rows[i] || [];
    for (var c = 0; c < row.length; c += 1) {
      var text = toText(row[c]).trim();
      if (!text) continue;
      var n = normalize(text);
      if (n.indexOf("SEMAINE") >= 0 || n.indexOf("DU LUNDI") >= 0) {
        return text;
      }
    }
  }
  return fallback;
}

function findDateContextLabel(rows, headerRow, fallback) {
  var best = "";

  for (var i = headerRow; i >= Math.max(0, headerRow - 6); i -= 1) {
    var row = rows[i] || [];
    for (var c = 0; c < row.length; c += 1) {
      var text = toText(row[c]).trim();
      if (!text) continue;
      var n = normalize(text);
      if (containsFrenchWeekday(n) && extractMonthNumber(n)) {
        return text;
      }
      if (!best && (extractMonthNumber(n) || /20\d{2}/.test(n))) {
        best = text;
      }
    }
  }

  return best || fallback;
}

function extractAssignments(rows, header, nextHeaderRow, fileName, fileOrder, sheetName, sheetOrder, weekLabel, dateContextLabel, weekIndex, weekMeta) {
  var out = [];
  var currentSection = "";
  var currentRole = "";
  var currentObservation = "";

  for (var r = header.row + 1; r < nextHeaderRow; r += 1) {
    var row = rows[r] || [];
    var sectionCell = toText(row[0]).trim();
    var roleCell = toText(row[1]).trim();
    var observationText = "";

    if (header.observationCol >= 0) {
      observationText = toText(row[header.observationCol]).trim();
      if (observationText) currentObservation = observationText;
    }

    if (sectionCell) currentSection = sectionCell;
    if (roleCell) currentRole = roleCell;

    var hasAnyAssignment = false;

    if (observationText) {
      var obsPeople = extractPeopleFromObservation(observationText);
      for (var op = 0; op < obsPeople.length; op += 1) {
        out.push({
          rawPerson: obsPeople[op],
          observationOnly: true,
          dayLabel: "",
          dayName: "",
          dayOrder: 99,
          dayNumber: 99,
          weekLabel: weekLabel,
          dateContextLabel: dateContextLabel,
          weekOrder: weekIndex,
          sortYear: weekMeta.year,
          sortMonth: weekMeta.month,
          fileOrder: fileOrder,
          sheetName: sheetName,
          sheetOrder: sheetOrder,
          fileName: fileName,
          section: currentSection || "",
          role: currentRole || "",
          task: "",
          observation: observationText,
        });
      }
    }

    for (var i = 0; i < header.dayColumns.length; i += 1) {
      var dayCol = header.dayColumns[i];
      var rawCell = toText(row[dayCol.col]).trim();
      if (!rawCell) continue;

      var people = splitPeople(rawCell);
      if (!people.length) continue;
      hasAnyAssignment = true;

      for (var p = 0; p < people.length; p += 1) {
        out.push({
          rawPerson: people[p],
          observationOnly: false,
          dayLabel: dayCol.dayLabel,
          dayName: dayCol.dayName,
          dayOrder: dayCol.dayOrder,
          dayNumber: extractDayNumber(dayCol.dayLabel),
          weekLabel: weekLabel,
          dateContextLabel: dateContextLabel,
          weekOrder: weekIndex,
          sortYear: weekMeta.year,
          sortMonth: weekMeta.month,
          fileOrder: fileOrder,
          sheetName: sheetName,
          sheetOrder: sheetOrder,
          fileName: fileName,
          section: currentSection || "",
          role: currentRole || "",
          task: buildTask(currentSection, currentRole),
          observation: currentObservation && observationMatchesPerson(currentObservation, people[p]) ? currentObservation : "",
        });
      }
    }

    if (!hasAnyAssignment && isSeparatorRow(row)) {
      currentRole = "";
    }
  }

  return out;
}

function mapRecords(records, nameMap) {
  var map = nameMap || {};
  return records
    .map(function (rec) {
      var mapped = applyNameMap(rec.rawPerson, map);
      if (!mapped) return null;
      var copy = {};
      Object.keys(rec).forEach(function (key) {
        copy[key] = rec[key];
      });
      copy.person = mapped;
      return copy;
    })
    .filter(Boolean);
}

function applyNameMap(rawName, nameMap) {
  var key = normalize(rawName);
  return nameMap[key] || rawName;
}

function extractMappedPeople(records) {
  var unique = {};
  records.forEach(function (rec) {
    unique[rec.person] = true;
  });
  return sortStrings(Object.keys(unique));
}

function getVisiblePersonRecords(records, personName, selectedTasks) {
  return records
    .filter(function (rec) {
      return rec.person === personName && isTaskEnabled(rec, selectedTasks || null);
    })
    .sort(sortRecords);
}

function isTaskEnabled(record, selectedTasks) {
  if (!selectedTasks) return true;
  if (!record.task) return true;
  return selectedTasks[record.task] !== false;
}

function sortRecords(a, b) {
  if (a.fileOrder !== b.fileOrder) return a.fileOrder - b.fileOrder;
  if (a.sheetOrder !== b.sheetOrder) return a.sheetOrder - b.sheetOrder;
  if (a.weekOrder !== b.weekOrder) return a.weekOrder - b.weekOrder;
  if (a.dayOrder !== b.dayOrder) return a.dayOrder - b.dayOrder;
  if (a.dayNumber !== b.dayNumber) return a.dayNumber - b.dayNumber;
  return toText(a.task).localeCompare(toText(b.task), "fr", { sensitivity: "base" });
}

function buildCalendarDayGroups(personRecords) {
  var weekObservations = {};
  var dayMap = {};
  var skipped = {};

  personRecords.forEach(function (rec) {
    var weekKey = rec.fileName + "|" + rec.sheetName + "|" + rec.weekLabel;
    if (!weekObservations[weekKey]) weekObservations[weekKey] = { list: [], set: {} };
    if (rec.observation && !weekObservations[weekKey].set[rec.observation]) {
      weekObservations[weekKey].set[rec.observation] = true;
      weekObservations[weekKey].list.push(rec.observation);
    }
  });

  personRecords.forEach(function (rec) {
    if (rec.observationOnly) return;
    var dayKey = rec.fileName + "|" + rec.sheetName + "|" + rec.weekLabel + "|" + rec.dayLabel;
    if (!dayMap[dayKey]) {
      var date = resolveRecordDate(rec);
      if (!date) {
        skipped[dayKey] = true;
        return;
      }
      var weekKey = rec.fileName + "|" + rec.sheetName + "|" + rec.weekLabel;
      dayMap[dayKey] = {
        date: date,
        dayLabel: rec.dayLabel,
        fileName: rec.fileName,
        sheetName: rec.sheetName,
        weekLabel: rec.weekLabel,
        tasks: [],
        taskSet: {},
        observations: weekObservations[weekKey] ? weekObservations[weekKey].list.slice() : [],
      };
    }
    if (rec.task && !dayMap[dayKey].taskSet[rec.task]) {
      dayMap[dayKey].taskSet[rec.task] = true;
      dayMap[dayKey].tasks.push(rec.task);
    }
  });

  var days = Object.keys(dayMap)
    .map(function (key) {
      return dayMap[key];
    })
    .sort(function (a, b) {
      return a.date.getTime() - b.date.getTime();
    });

  return { days: days, skippedCount: Object.keys(skipped).length };
}

function buildIcsText(personName, days) {
  var lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//planner//planning export//FR", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"];
  var stamp = formatUtcTimestamp(new Date());

  days.forEach(function (day) {
    var uid = slugifyFileName(personName) + "-" + formatDateKey(day.date) + "-" + slugifyFileName(day.fileName + "-" + day.sheetName);
    lines.push("BEGIN:VEVENT");
    lines.push("UID:" + uid + "@planner");
    lines.push("DTSTAMP:" + stamp);
    lines.push("SUMMARY:" + escapeIcsText(buildIcsSummary(personName, day)));
    lines.push("DESCRIPTION:" + escapeIcsText(buildIcsDescription(day)));
    lines.push("DTSTART;VALUE=DATE:" + formatDateKey(day.date));
    lines.push("DTEND;VALUE=DATE:" + formatDateKey(addUtcDays(day.date, 1)));
    lines.push("END:VEVENT");
  });

  lines.push("END:VCALENDAR");
  return foldIcsLines(lines).join("\r\n") + "\r\n";
}

function buildIcsSummary(personName, day) {
  return personName + " - " + (day.tasks.length ? day.tasks.join(" | ") : "Planning");
}

function buildIcsDescription(day) {
  var lines = [];
  if (day.tasks.length) lines.push("Taches: " + day.tasks.join(" | "));
  if (day.observations.length) lines.push("Observations: " + day.observations.join(" | "));
  lines.push("Jour: " + day.dayLabel);
  lines.push("Source: " + day.fileName + " / " + day.sheetName);
  lines.push("Semaine: " + day.weekLabel);
  return lines.join("\n");
}

function resolveRecordDate(rec) {
  var explicitDate = parseExplicitDayLabel(rec.dayLabel, rec.dateContextLabel || rec.weekLabel || "", rec.sortYear);
  if (explicitDate) return explicitDate;

  var dayNumber = rec.dayNumber || extractDayNumber(rec.dayLabel);
  if (!dayNumber) return null;

  var weekRange = parseWeekRange(rec.dateContextLabel || rec.weekLabel || "");
  var weekdayOrder = getWeekdayOrder(rec.dayName || rec.dayLabel || "");
  if (weekRange && weekdayOrder >= 0) {
    var startDate = new Date(Date.UTC(weekRange.year, weekRange.startMonth - 1, weekRange.startDay));
    return addUtcDays(startDate, weekdayOrder);
  }

  var context = [rec.dateContextLabel, rec.weekLabel, rec.sheetName, rec.fileName].join(" ");
  var month = extractMonthNumber(context) || rec.sortMonth;
  var year = extractYearNumber(context) || rec.sortYear || new Date().getFullYear();
  if (!month || !year) return null;
  return new Date(Date.UTC(year, month - 1, dayNumber));
}

function extractMonthNumber(text) {
  var normalized = normalize(text);
  var months = { JANVIER: 1, FEVRIER: 2, MARS: 3, AVRIL: 4, MAI: 5, JUIN: 6, JUILLET: 7, AOUT: 8, SEPTEMBRE: 9, OCTOBRE: 10, NOVEMBRE: 11, DECEMBRE: 12 };
  var tokens = normalized.match(/[A-Z]+/g) || [];
  for (var i = 0; i < tokens.length; i += 1) {
    if (months[tokens[i]]) return months[tokens[i]];
  }
  return 0;
}

function extractYearNumber(text) {
  var match = toText(text).match(/(20\d{2})/);
  return match ? parseInt(match[1], 10) : 0;
}

function foldIcsLines(lines) {
  var folded = [];
  lines.forEach(function (line) {
    var text = toText(line);
    while (text.length > 75) {
      folded.push(text.slice(0, 75));
      text = " " + text.slice(75);
    }
    folded.push(text);
  });
  return folded;
}

function formatDateKey(date) {
  return date.getUTCFullYear().toString() + pad2(date.getUTCMonth() + 1) + pad2(date.getUTCDate());
}

function formatUtcTimestamp(date) {
  return formatDateKey(date) + "T" + pad2(date.getUTCHours()) + pad2(date.getUTCMinutes()) + pad2(date.getUTCSeconds()) + "Z";
}

function addUtcDays(date, days) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate() + days));
}

function pad2(value) {
  return value < 10 ? "0" + value : String(value);
}

function slugifyFileName(text) {
  return normalize(text).replace(/[^A-Z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase();
}

function escapeIcsText(text) {
  return toText(text).replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\r?\n/g, "\\n");
}

function containsFrenchWeekday(text) {
  var normalized = normalize(text);
  for (var i = 0; i < dayKeys.length; i += 1) {
    if (normalized.indexOf(dayKeys[i]) >= 0) return true;
  }
  return false;
}

function getWeekdayOrder(dayName) {
  var normalized = normalize(dayName);
  for (var i = 0; i < dayKeys.length; i += 1) {
    if (normalized.indexOf(dayKeys[i]) === 0) return i;
  }
  return -1;
}

function parseExplicitDayLabel(dayLabel, contextLabel, fallbackYear) {
  var label = normalize(dayLabel);
  var year = extractYearNumber(contextLabel) || fallbackYear || 0;
  if (!year) return null;

  var slashMatch = label.match(/(\d{1,2})\s*\/\s*(\d{1,2})/);
  if (slashMatch) {
    return new Date(Date.UTC(year, parseInt(slashMatch[2], 10) - 1, parseInt(slashMatch[1], 10)));
  }

  return null;
}

function parseWeekRange(text) {
  var normalized = normalize(text).replace(/(JUILLET|JUIN|AOUT|DECEMBRE|NOVEMBRE|OCTOBRE|SEPTEMBRE|FEVRIER|JANVIER|MARS|AVRIL|MAI)(20\d{2})/, "$1 $2");
  var regex = /DU\s+[A-Z]+\s+(\d{1,2})(?:ER)?(?:\s+([A-Z]+))?\s+AU\s+(?:[A-Z]+\s+)?(\d{1,2})(?:ER)?(?:\s+([A-Z]+))?\s*(20\d{2})/;
  var match = normalized.match(regex);
  if (!match) return null;

  var startDay = parseInt(match[1], 10);
  var startMonth = extractMonthNumber(match[2] || "") || 0;
  var endDay = parseInt(match[3], 10);
  var endMonth = extractMonthNumber(match[4] || "") || 0;
  var year = parseInt(match[5], 10);

  if (!startMonth && endMonth) startMonth = endMonth;
  if (!startMonth || !year || !startDay || !endDay) return null;

  return {
    startDay: startDay,
    startMonth: startMonth,
    endDay: endDay,
    endMonth: endMonth || startMonth,
    year: year,
  };
}

function splitPeople(rawCell) {
  var original = toText(rawCell).replace(/\r/g, "");
  var cleaned = original.replace(/\s+/g, " ").trim();
  if (!cleaned) return [];
  var parts = original.split(/\s{2,}|\s\/\s|;|,|&|\n/).map(sanitizePerson).filter(Boolean);
  if (!parts.length) {
    var fallback = sanitizePerson(cleaned);
    return fallback ? [fallback] : [];
  }
  var unique = {};
  var out = [];
  parts.forEach(function (p) {
    if (!unique[p]) {
      unique[p] = true;
      out.push(p);
    }
  });
  return out;
}

function extractPeopleFromObservation(observation) {
  var text = toText(observation);
  if (!text) return [];
  var source = text;
  var colonIndex = source.indexOf(":");
  if (colonIndex >= 0) source = source.slice(colonIndex + 1);
  source = source.replace(/\bet\b/gi, ",");
  var candidates = source.split(/[;,/]/).map(sanitizePerson).filter(Boolean);
  var unique = {};
  var out = [];
  candidates.forEach(function (name) {
    if (!unique[name]) {
      unique[name] = true;
      out.push(name);
    }
  });
  return out;
}

function observationMatchesPerson(observation, person) {
  var obs = normalize(observation);
  var who = normalize(person);
  if (!obs || !who) return false;
  if (obs.indexOf(who) >= 0) return true;
  var tokens = who.split(/[^A-Z0-9]+/).filter(function (t) {
    return t.length >= 3;
  });
  for (var i = 0; i < tokens.length; i += 1) {
    if (obs.indexOf(tokens[i]) >= 0) return true;
  }
  return false;
}

function sanitizePerson(text) {
  var value = toText(text).replace(/\s+/g, " ").trim();
  if (!value) return "";
  value = value.replace(/\s+CP$/i, "").trim();
  value = value.replace(/\s+enr\b.*$/i, "").trim();
  value = value.replace(/\?+$/g, "").trim();
  value = value.replace(/\s+Rep$/i, "").trim();
  value = value.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
  var lowerIndex = value.search(/\s+[a-z][^A-Z]*/);
  if (lowerIndex > 0) {
    var shortened = value.slice(0, lowerIndex).trim();
    if (shortened) value = shortened;
  }
  if (looksLikeNonPerson(value)) return "";
  if (!/[A-Za-zÀ-ÖØ-öø-ÿ]/.test(value)) return "";
  return value.toLocaleUpperCase("fr-FR");
}

function looksLikeNonPerson(value) {
  var normalized = normalize(value);
  if (!normalized) return true;
  if (normalized.indexOf("PLANNING") >= 0) return true;
  if (normalized.indexOf("SEMAINE") >= 0) return true;
  if (normalized.indexOf("OBSERVATION") >= 0) return true;
  if (containsFrenchWeekday(normalized)) return true;
  return false;
}

function buildTask(section, role) {
  var s = toText(section).trim();
  var r = toText(role).trim();
  return s && r ? s + " - " + r : s || r || "";
}

function isSeparatorRow(row) {
  if (!row || !row.length) return true;
  for (var i = 0; i < row.length; i += 1) {
    if (toText(row[i]).trim()) return false;
  }
  return true;
}

function extractDayNumber(dayLabel) {
  var m = toText(dayLabel).match(/(\d{1,2})/);
  return m ? parseInt(m[1], 10) : 0;
}

function deriveWeekMeta(weekLabel) {
  return { year: extractYearNumber(weekLabel), month: extractMonthNumber(weekLabel) };
}

function normalize(value) {
  return toText(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/\s+/g, " ").trim();
}

function sortStrings(values) {
  return values.sort(function (a, b) {
    return a.localeCompare(b, "fr", { sensitivity: "base" });
  });
}

function toText(value) {
  return value === undefined || value === null ? "" : String(value);
}

var exported = {
  dayKeys: dayKeys,
  parseWorkbook: parseWorkbook,
  mapRecords: mapRecords,
  extractMappedPeople: extractMappedPeople,
  getVisiblePersonRecords: getVisiblePersonRecords,
  buildCalendarDayGroups: buildCalendarDayGroups,
  buildIcsText: buildIcsText,
  resolveRecordDate: resolveRecordDate,
  extractMonthNumber: extractMonthNumber,
  extractYearNumber: extractYearNumber,
  extractDayNumber: extractDayNumber,
  deriveWeekMeta: deriveWeekMeta,
  detectHeaderRows: detectHeaderRows,
  findWeekLabel: findWeekLabel,
  findDateContextLabel: findDateContextLabel,
  extractPeopleFromObservation: extractPeopleFromObservation,
  observationMatchesPerson: observationMatchesPerson,
  splitPeople: splitPeople,
  sanitizePerson: sanitizePerson,
  normalize: normalize,
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = exported;
}

if (typeof window !== "undefined") {
  window.PlanningCore = exported;
}
