(function () {
  var state = {
    records: [],
    people: [],
    selectedPerson: "",
    showTasks: false,
  };

  var dayKeys = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"];

  var fileInput = document.getElementById("fileInput");
  var parseBtn = document.getElementById("parseBtn");
  var statusEl = document.getElementById("status");
  var controlsEl = document.getElementById("controls");
  var summaryEl = document.getElementById("summary");
  var personViewEl = document.getElementById("personView");
  var personSearchEl = document.getElementById("personSearch");
  var personSelectEl = document.getElementById("personSelect");
  var showTasksEl = document.getElementById("showTasks");

  parseBtn.addEventListener("click", onParseClick);
  personSearchEl.addEventListener("input", onSearchInput);
  personSelectEl.addEventListener("change", onPersonSelect);
  showTasksEl.addEventListener("change", function () {
    state.showTasks = showTasksEl.checked;
    renderPerson();
  });

  function onParseClick() {
    if (!window.XLSX) {
      setStatus("La librairie XLSX n'est pas chargee.", true);
      return;
    }

    var file = fileInput.files && fileInput.files[0];
    if (!file) {
      setStatus("Selectionnez un fichier .xlsx.", true);
      return;
    }

    setStatus("Lecture du fichier...", false);
    var reader = new FileReader();
    reader.onload = function (evt) {
      try {
        var data = evt.target.result;
        var workbook = XLSX.read(data, { type: "array", cellDates: false });
        var parsed = parseWorkbook(workbook);
        state.records = parsed;
        state.people = extractPeople(parsed);
        state.selectedPerson = state.people[0] || "";
        state.showTasks = showTasksEl.checked;

        if (!state.records.length) {
          controlsEl.classList.add("hidden");
          summaryEl.classList.add("hidden");
          personViewEl.innerHTML = "";
          setStatus("Aucune affectation detectee dans ce fichier.", true);
          return;
        }

        setStatus("Planning genere avec succes.", false);
        renderControls();
        renderSummary();
        renderPerson();
      } catch (err) {
        console.error(err);
        controlsEl.classList.add("hidden");
        summaryEl.classList.add("hidden");
        personViewEl.innerHTML = "";
        setStatus("Erreur pendant l'analyse du fichier.", true);
      }
    };
    reader.onerror = function () {
      setStatus("Impossible de lire le fichier.", true);
    };
    reader.readAsArrayBuffer(file);
  }

  function parseWorkbook(workbook) {
    var records = [];
    var weekIndex = 0;

    workbook.SheetNames.forEach(function (sheetName) {
      var sheet = workbook.Sheets[sheetName];
      var rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: "",
      });

      var headers = detectHeaderRows(rows);
      headers.forEach(function (header, idx) {
        weekIndex += 1;
        var nextHeaderRow = idx < headers.length - 1 ? headers[idx + 1].row : rows.length;
        var weekLabel = findWeekLabel(rows, header.row, sheetName);
        records = records.concat(
          extractAssignments(rows, header, nextHeaderRow, sheetName, weekLabel, weekIndex)
        );
      });
    });

    return records;
  }

  function detectHeaderRows(rows) {
    var headers = [];

    for (var r = 0; r < rows.length; r += 1) {
      var row = rows[r] || [];
      var dayColumns = [];

      for (var c = 0; c < row.length; c += 1) {
        var value = toText(row[c]);
        if (!value) continue;
        var n = normalize(value);
        for (var d = 0; d < dayKeys.length; d += 1) {
          if (n.indexOf(dayKeys[d]) === 0) {
            dayColumns.push({
              col: c,
              dayName: dayKeys[d],
              dayLabel: value.trim(),
              dayOrder: d,
            });
            break;
          }
        }
      }

      if (dayColumns.length >= 5) {
        dayColumns.sort(function (a, b) {
          return a.col - b.col;
        });
        headers.push({ row: r, dayColumns: dayColumns });
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

  function extractAssignments(rows, header, nextHeaderRow, sheetName, weekLabel, weekIndex) {
    var out = [];
    var currentSection = "";
    var currentRole = "";

    for (var r = header.row + 1; r < nextHeaderRow; r += 1) {
      var row = rows[r] || [];
      var sectionCell = toText(row[0]).trim();
      var roleCell = toText(row[1]).trim();

      if (sectionCell) currentSection = sectionCell;
      if (roleCell) currentRole = roleCell;

      var hasAnyAssignment = false;

      for (var i = 0; i < header.dayColumns.length; i += 1) {
        var dayCol = header.dayColumns[i];
        var rawCell = toText(row[dayCol.col]).trim();
        if (!rawCell) continue;

        var people = splitPeople(rawCell);
        if (!people.length) continue;
        hasAnyAssignment = true;

        for (var p = 0; p < people.length; p += 1) {
          out.push({
            person: people[p],
            dayLabel: dayCol.dayLabel,
            dayName: dayCol.dayName,
            dayOrder: dayCol.dayOrder,
            dayNumber: extractDayNumber(dayCol.dayLabel),
            weekLabel: weekLabel,
            weekOrder: weekIndex,
            sheetName: sheetName,
            section: currentSection || "",
            role: currentRole || "",
            task: buildTask(currentSection, currentRole),
          });
        }
      }

      if (!hasAnyAssignment && isSeparatorRow(row)) {
        currentRole = "";
      }
    }

    return out;
  }

  function extractPeople(records) {
    var unique = {};
    records.forEach(function (rec) {
      unique[rec.person] = true;
    });
    return Object.keys(unique).sort(function (a, b) {
      return a.localeCompare(b, "fr", { sensitivity: "base" });
    });
  }

  function renderControls() {
    controlsEl.classList.remove("hidden");
    personSelectEl.innerHTML = "";

    state.people.forEach(function (name) {
      var opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      personSelectEl.appendChild(opt);
    });

    personSelectEl.value = state.selectedPerson;
    personSearchEl.value = state.selectedPerson;
  }

  function renderSummary() {
    summaryEl.classList.remove("hidden");

    var daysCountByPerson = {};
    state.records.forEach(function (rec) {
      var key = rec.person + "|" + rec.weekOrder + "|" + rec.dayName + "|" + rec.dayLabel;
      daysCountByPerson[key] = true;
    });

    var totalAssignments = state.records.length;
    var totalDistinctDays = Object.keys(daysCountByPerson).length;

    summaryEl.innerHTML =
      '<div class="summary-grid">' +
      '<div class="kpi"><strong>' + state.people.length + "</strong><span>personnes</span></div>" +
      '<div class="kpi"><strong>' + totalDistinctDays + "</strong><span>jours planifies</span></div>" +
      '<div class="kpi"><strong>' + totalAssignments + "</strong><span>affectations</span></div>" +
      "</div>";
  }

  function onSearchInput() {
    var q = personSearchEl.value.trim().toLowerCase();
    if (!q) {
      return;
    }

    var found = state.people.find(function (name) {
      return name.toLowerCase().indexOf(q) >= 0;
    });

    if (found) {
      state.selectedPerson = found;
      personSelectEl.value = found;
      renderPerson();
    }
  }

  function onPersonSelect() {
    state.selectedPerson = personSelectEl.value;
    personSearchEl.value = state.selectedPerson;
    renderPerson();
  }

  function renderPerson() {
    var name = state.selectedPerson;
    if (!name) {
      personViewEl.innerHTML = "";
      return;
    }

    var personRecords = state.records
      .filter(function (rec) {
        return rec.person === name;
      })
      .sort(sortRecords);

    if (!personRecords.length) {
      personViewEl.innerHTML = '<article class="panel"><p>Aucun planning pour cette personne.</p></article>';
      return;
    }

    var timeline = buildTimeline(personRecords);
    var html = '<article class="panel">';
    html += "<h2>" + escapeHtml(name) + "</h2>";

    timeline.forEach(function (week) {
      html += '<section class="week-block">';
      html += '<h3 class="week-title">' + escapeHtml(week.weekLabel) + "</h3>";
      html += '<ul class="day-list">';

      week.days.forEach(function (day) {
        html += '<li class="day-item">';
        html += '<div class="day-main"><span class="day-chip">' + escapeHtml(day.dayLabel) + "</span></div>";

        if (state.showTasks && day.tasks.length) {
          html += '<ul class="task-list">';
          day.tasks.forEach(function (task) {
            html += "<li>" + escapeHtml(task) + "</li>";
          });
          html += "</ul>";
        }

        html += "</li>";
      });

      html += "</ul>";
      html += "</section>";
    });

    html += "</article>";
    personViewEl.innerHTML = html;
  }

  function buildTimeline(records) {
    var weeks = [];
    var weekMap = {};

    records.forEach(function (rec) {
      var wkKey = rec.weekOrder + "|" + rec.weekLabel;
      if (!weekMap[wkKey]) {
        weekMap[wkKey] = {
          weekOrder: rec.weekOrder,
          weekLabel: rec.weekLabel,
          days: [],
          dayMap: {},
        };
        weeks.push(weekMap[wkKey]);
      }

      var dayKey = rec.dayOrder + "|" + rec.dayLabel;
      if (!weekMap[wkKey].dayMap[dayKey]) {
        weekMap[wkKey].dayMap[dayKey] = {
          dayOrder: rec.dayOrder,
          dayLabel: rec.dayLabel,
          tasks: [],
          taskSet: {},
        };
        weekMap[wkKey].days.push(weekMap[wkKey].dayMap[dayKey]);
      }

      if (rec.task && !weekMap[wkKey].dayMap[dayKey].taskSet[rec.task]) {
        weekMap[wkKey].dayMap[dayKey].taskSet[rec.task] = true;
        weekMap[wkKey].dayMap[dayKey].tasks.push(rec.task);
      }
    });

    weeks.sort(function (a, b) {
      return a.weekOrder - b.weekOrder;
    });

    weeks.forEach(function (w) {
      w.days.sort(function (a, b) {
        if (a.dayOrder !== b.dayOrder) return a.dayOrder - b.dayOrder;
        var an = extractDayNumber(a.dayLabel);
        var bn = extractDayNumber(b.dayLabel);
        return an - bn;
      });
    });

    return weeks;
  }

  function sortRecords(a, b) {
    if (a.weekOrder !== b.weekOrder) return a.weekOrder - b.weekOrder;
    if (a.dayOrder !== b.dayOrder) return a.dayOrder - b.dayOrder;
    if (a.dayNumber !== b.dayNumber) return a.dayNumber - b.dayNumber;
    return a.task.localeCompare(b.task, "fr", { sensitivity: "base" });
  }

  function splitPeople(rawCell) {
    var original = toText(rawCell).replace(/\r/g, "");
    var cleaned = original.replace(/\s+/g, " ").trim();
    if (!cleaned) return [];

    var parts = original
      .split(/\s{2,}|\s\/\s|;|,|&|\n/)
      .map(function (s) {
        return sanitizePerson(s);
      })
      .filter(Boolean);

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

  function sanitizePerson(text) {
    var value = toText(text).replace(/\s+/g, " ").trim();
    if (!value) return "";

    value = value.replace(/\s+CP$/i, "").trim();
    value = value.replace(/\s+enr\b.*$/i, "").trim();

    var lowerIndex = value.search(/\s+[a-z][^A-Z]*/);
    if (lowerIndex > 0) {
      var shortened = value.slice(0, lowerIndex).trim();
      if (shortened) value = shortened;
    }

    if (!/[A-Za-zÀ-ÖØ-öø-ÿ]/.test(value)) return "";
    return value;
  }

  function buildTask(section, role) {
    var s = toText(section).trim();
    var r = toText(role).trim();
    if (s && r) return s + " - " + r;
    return s || r || "";
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

  function normalize(value) {
    return toText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function toText(value) {
    if (value === undefined || value === null) return "";
    return String(value);
  }

  function setStatus(message, isError) {
    statusEl.textContent = message;
    statusEl.classList.toggle("error", !!isError);
  }

  function escapeHtml(value) {
    return toText(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
})();
