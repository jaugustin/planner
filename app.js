(function () {
  var core = window.PlanningCore;
  var STORAGE_KEYS = {
    nameMap: "planning_name_map_auto_v2",
    selectedPerson: "planning_selected_person_v1",
    showTasks: "planning_show_tasks_v1",
  };

  var dayKeys = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"];

  var state = {
    baseRecords: [],
    records: [],
    people: [],
    rawPeople: [],
    taskList: [],
    selectedTasks: {},
    selectedPerson: loadString(STORAGE_KEYS.selectedPerson),
    showTasks: loadBool(STORAGE_KEYS.showTasks),
    nameMap: loadNameMap(),
  };

  var fileInput = document.getElementById("fileInput");
  var parseBtn = document.getElementById("parseBtn");
  var statusEl = document.getElementById("status");
  var controlsEl = document.getElementById("controls");
  var summaryEl = document.getElementById("summary");
  var exportPanelEl = document.getElementById("exportPanel");
  var personViewEl = document.getElementById("personView");
  var personSearchEl = document.getElementById("personSearch");
  var personSelectEl = document.getElementById("personSelect");
  var showTasksEl = document.getElementById("showTasks");
  var taskChecklistEl = document.getElementById("taskChecklist");
  var exportIcsBtnEl = document.getElementById("exportIcsBtn");
  var shareIcsBtnEl = document.getElementById("shareIcsBtn");

  showTasksEl.checked = state.showTasks;

  parseBtn.addEventListener("click", onParseClick);
  personSearchEl.addEventListener("input", onSearchInput);
  personSelectEl.addEventListener("change", onPersonSelect);
  showTasksEl.addEventListener("change", function () {
    state.showTasks = showTasksEl.checked;
    saveString(STORAGE_KEYS.showTasks, state.showTasks ? "1" : "0");
    renderPerson();
  });
  taskChecklistEl.addEventListener("change", onTaskChecklistChange);
  exportIcsBtnEl.addEventListener("click", onExportIcs);
  shareIcsBtnEl.addEventListener("click", onShareIcs);

  function onParseClick() {
    if (!window.XLSX) {
      setStatus("La librairie XLSX n'est pas chargee.", true);
      return;
    }

    var files = Array.prototype.slice.call(fileInput.files || []);
    if (!files.length) {
      setStatus("Selectionnez un ou plusieurs fichiers .xlsx.", true);
      return;
    }

    setStatus("Lecture de " + files.length + " fichier(s)...", false);

    Promise.all(
      files.map(function (file, index) {
        return readWorkbookFromFile(file, index);
      })
    )
      .then(function (workbooks) {
        var merged = [];
        workbooks.forEach(function (item) {
          merged = merged.concat(parseWorkbook(item.workbook, item.fileName, item.fileOrder));
        });

        state.baseRecords = merged;
        if (!state.baseRecords.length) {
          state.records = [];
          state.people = [];
          controlsEl.classList.add("hidden");
          summaryEl.classList.add("hidden");
          exportPanelEl.classList.add("hidden");
          personViewEl.innerHTML = "";
          setStatus("Aucune affectation detectee dans les fichiers importes.", true);
          return;
        }

        refreshDerivedState();
        setStatus(
          "Planning fusionne genere: " +
            files.length +
            " fichier(s), " +
            state.people.length +
            " personne(s).",
          false
        );
      })
      .catch(function (err) {
        console.error(err);
        controlsEl.classList.add("hidden");
        summaryEl.classList.add("hidden");
        exportPanelEl.classList.add("hidden");
        personViewEl.innerHTML = "";
        setStatus("Erreur pendant l'analyse des fichiers.", true);
      });
  }

  function refreshDerivedState() {
    updateAutoNameMap(state.baseRecords);

    state.records = core.mapRecords(state.baseRecords, state.nameMap);
    state.people = core.extractMappedPeople(state.records);
    state.rawPeople = extractRawPeople(state.baseRecords);
    refreshTaskFilter(state.records);

    if (!state.selectedPerson || state.people.indexOf(state.selectedPerson) < 0) {
      state.selectedPerson = state.people[0] || "";
      saveString(STORAGE_KEYS.selectedPerson, state.selectedPerson || "");
    }

    renderControls();
    renderSummary();
    renderPerson();
  }

  function readWorkbookFromFile(file, fileOrder) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (evt) {
        try {
          var workbook = XLSX.read(evt.target.result, { type: "array", cellDates: false });
          resolve({ fileName: file.name, fileOrder: fileOrder, workbook: workbook });
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  function parseWorkbook(workbook, fileName, fileOrder) {
    return core.parseWorkbook(workbook, fileName, fileOrder);
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
        if (core.normalize(text) && core.extractMonthNumber(text) && dayKeys.some(function (dayKey) { return core.normalize(text).indexOf(dayKey) >= 0; })) {
          return text;
        }
        if (!best && (core.extractMonthNumber(text) || /20\d{2}/.test(n))) {
          best = text;
        }
      }
    }

    return best || fallback;
  }

  function extractAssignments(
    rows,
    header,
    nextHeaderRow,
    fileName,
    fileOrder,
    sheetName,
    sheetOrder,
    weekLabel,
    dateContextLabel,
    weekIndex,
    weekMeta
  ) {
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
        if (observationText) {
          currentObservation = observationText;
        }
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
            observation:
              currentObservation && observationMatchesPerson(currentObservation, people[p])
                ? currentObservation
                : "",
          });
        }
      }

      if (!hasAnyAssignment && isSeparatorRow(row)) {
        currentRole = "";
      }
    }

    return out;
  }

  function extractMappedPeople(records) {
    var unique = {};
    records.forEach(function (rec) {
      unique[rec.person] = true;
    });
    return sortStrings(Object.keys(unique));
  }

  function extractRawPeople(records) {
    var unique = {};
    records.forEach(function (rec) {
      unique[rec.rawPerson] = true;
    });
    return sortStrings(Object.keys(unique));
  }

  function refreshTaskFilter(records) {
    var unique = {};
    records.forEach(function (rec) {
      if (rec.task) unique[rec.task] = true;
    });
    state.taskList = sortStrings(Object.keys(unique));

    var next = {};
    state.taskList.forEach(function (task) {
      if (Object.prototype.hasOwnProperty.call(state.selectedTasks, task)) {
        next[task] = state.selectedTasks[task];
      } else {
        next[task] = true;
      }
    });
    state.selectedTasks = next;
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
    renderTaskChecklist();
  }

  function renderTaskChecklist() {
    if (!state.taskList.length) {
      taskChecklistEl.innerHTML = '<p class="hint">Aucune tache detectee.</p>';
      return;
    }

    var html = "";
    state.taskList.forEach(function (task) {
      var checked = state.selectedTasks[task] !== false ? " checked" : "";
      html +=
        '<label class="task-item"><input type="checkbox" data-task="' +
        escapeHtml(task) +
        '"' +
        checked +
        " />" +
        escapeHtml(task) +
        "</label>";
    });
    taskChecklistEl.innerHTML = html;
  }

  function renderSummary() {
    summaryEl.classList.remove("hidden");

    var enabled = state.records.filter(isTaskEnabled);
    var daysCountByPerson = {};
    enabled.forEach(function (rec) {
      if (rec.observationOnly) return;
      var key = rec.person + "|" + rec.fileName + "|" + rec.sheetName + "|" + rec.weekLabel + "|" + rec.dayLabel;
      daysCountByPerson[key] = true;
    });

    var globalDaysSelectedPerson = countGlobalDaysForPerson(enabled, state.selectedPerson);
    var selectedLabel = state.selectedPerson ? "(" + state.selectedPerson + ")" : "";

    summaryEl.innerHTML =
      '<div class="summary-grid">' +
      '<div class="kpi"><strong>' + state.people.length + "</strong><span>personnes</span></div>" +
      '<div class="kpi"><strong>' + Object.keys(daysCountByPerson).length + "</strong><span>jours planifies</span></div>" +
      '<div class="kpi"><strong>' + enabled.length + "</strong><span>affectations</span></div>" +
      '<div class="kpi"><strong>' +
      globalDaysSelectedPerson +
      "</strong><span>jours globaux " +
      escapeHtml(selectedLabel) +
      "</span></div>" +
      "</div>";
  }

  function getVisiblePersonRecords(name) {
    return core.getVisiblePersonRecords(state.records, name, state.selectedTasks);
  }

  function renderExportPanel(personRecords) {
    var hasExportableDay = personRecords.some(function (rec) {
      return !rec.observationOnly;
    });

    if (!state.selectedPerson || !personRecords.length || !hasExportableDay) {
      exportPanelEl.classList.add("hidden");
      shareIcsBtnEl.disabled = true;
      return;
    }

    exportPanelEl.classList.remove("hidden");
    shareIcsBtnEl.disabled = !canShareIcsFile();
  }

  function renderPerson() {
    var name = state.selectedPerson;
    if (!name) {
      renderExportPanel([]);
      personViewEl.innerHTML = "";
      return;
    }

    var personRecords = getVisiblePersonRecords(name);
    renderExportPanel(personRecords);

    if (!personRecords.length) {
      personViewEl.innerHTML = '<article class="panel"><p>Aucun planning pour cette personne.</p></article>';
      return;
    }

    var timeline = buildTimeline(personRecords);
    var html = '<article class="panel">';
    html += "<h2>" + escapeHtml(name) + "</h2>";

    timeline.forEach(function (fileGroup) {
      html += '<section class="file-block">';
      html += '<h3 class="file-title">' + escapeHtml(fileGroup.fileName) + "</h3>";

      fileGroup.sheets.forEach(function (sheetGroup) {
        html += '<section class="sheet-block">';
        html += '<h4 class="sheet-title">' + escapeHtml(sheetGroup.sheetName) + "</h4>";

        sheetGroup.weeks.forEach(function (week) {
          html += '<section class="week-block">';
          html += '<h5 class="week-title">' + escapeHtml(week.weekLabel) + "</h5>";

          if (week.observations.length) {
            html += '<ul class="obs-list">';
            week.observations.forEach(function (observation) {
              html += "<li>Observation: " + escapeHtml(observation) + "</li>";
            });
            html += "</ul>";
          }

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

        html += "</section>";
      });

      html += "</section>";
    });

    html += "</article>";
    personViewEl.innerHTML = html;
  }

  function buildTimeline(records) {
    var files = [];
    var fileMap = {};

    records.forEach(function (rec) {
      var fileKey = rec.fileOrder + "|" + rec.fileName;
      if (!fileMap[fileKey]) {
        fileMap[fileKey] = {
          fileOrder: rec.fileOrder,
          fileName: rec.fileName,
          sheets: [],
          sheetMap: {},
        };
        files.push(fileMap[fileKey]);
      }

      var sheetKey = rec.sheetOrder + "|" + rec.sheetName;
      if (!fileMap[fileKey].sheetMap[sheetKey]) {
        fileMap[fileKey].sheetMap[sheetKey] = {
          sheetOrder: rec.sheetOrder,
          sheetName: rec.sheetName,
          weeks: [],
          weekMap: {},
        };
        fileMap[fileKey].sheets.push(fileMap[fileKey].sheetMap[sheetKey]);
      }

      var weekKey = rec.weekOrder + "|" + rec.weekLabel;
      if (!fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey]) {
        fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey] = {
          weekOrder: rec.weekOrder,
          weekLabel: rec.weekLabel,
          observations: [],
          observationSet: {},
          days: [],
          dayMap: {},
        };
        fileMap[fileKey].sheetMap[sheetKey].weeks.push(fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey]);
      }

      if (rec.observation) {
        var weekObj = fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey];
        if (!weekObj.observationSet[rec.observation]) {
          weekObj.observationSet[rec.observation] = true;
          weekObj.observations.push(rec.observation);
        }
      }

      if (rec.observationOnly) {
        return;
      }

      var dayKey = rec.dayOrder + "|" + rec.dayLabel;
      if (!fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey].dayMap[dayKey]) {
        fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey].dayMap[dayKey] = {
          dayOrder: rec.dayOrder,
          dayLabel: rec.dayLabel,
          tasks: [],
          taskSet: {},
        };
        fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey].days.push(
          fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey].dayMap[dayKey]
        );
      }

      if (rec.task && !fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey].dayMap[dayKey].taskSet[rec.task]) {
        fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey].dayMap[dayKey].taskSet[rec.task] = true;
        fileMap[fileKey].sheetMap[sheetKey].weekMap[weekKey].dayMap[dayKey].tasks.push(rec.task);
      }

    });

    files.sort(function (a, b) {
      if (a.fileOrder !== b.fileOrder) return a.fileOrder - b.fileOrder;
      return a.fileName.localeCompare(b.fileName, "fr", { sensitivity: "base" });
    });

    files.forEach(function (fileGroup) {
      fileGroup.sheets.sort(function (a, b) {
        if (a.sheetOrder !== b.sheetOrder) return a.sheetOrder - b.sheetOrder;
        return a.sheetName.localeCompare(b.sheetName, "fr", { sensitivity: "base" });
      });

      fileGroup.sheets.forEach(function (sheetGroup) {
        sheetGroup.weeks.sort(function (a, b) {
          if (a.weekOrder !== b.weekOrder) return a.weekOrder - b.weekOrder;
          return a.weekLabel.localeCompare(b.weekLabel, "fr", { sensitivity: "base" });
        });

        sheetGroup.weeks.forEach(function (week) {
          week.days.sort(function (a, b) {
            if (a.dayOrder !== b.dayOrder) return a.dayOrder - b.dayOrder;
            return extractDayNumber(a.dayLabel) - extractDayNumber(b.dayLabel);
          });
        });
      });
    });

    return files;
  }

  function sortRecords(a, b) {
    if (a.fileOrder !== b.fileOrder) return a.fileOrder - b.fileOrder;
    if (a.sheetOrder !== b.sheetOrder) return a.sheetOrder - b.sheetOrder;
    if (a.weekOrder !== b.weekOrder) return a.weekOrder - b.weekOrder;
    if (a.dayOrder !== b.dayOrder) return a.dayOrder - b.dayOrder;
    if (a.dayNumber !== b.dayNumber) return a.dayNumber - b.dayNumber;
    return a.task.localeCompare(b.task, "fr", { sensitivity: "base" });
  }

  function countGlobalDaysForPerson(records, personName) {
    if (!personName) return 0;
    var unique = {};
    records.forEach(function (rec) {
      if (rec.person !== personName) return;
      if (rec.observationOnly) return;
      var key = rec.fileName + "|" + rec.sheetName + "|" + rec.weekLabel + "|" + rec.dayLabel;
      unique[key] = true;
    });
    return Object.keys(unique).length;
  }

  function isTaskEnabled(record) {
    if (!record.task) return true;
    return state.selectedTasks[record.task] !== false;
  }

  function onSearchInput() {
    var q = personSearchEl.value.trim().toLowerCase();
    if (!q) return;

    var found = state.people.find(function (name) {
      return name.toLowerCase().indexOf(q) >= 0;
    });

    if (found) {
      state.selectedPerson = found;
      personSelectEl.value = found;
      saveString(STORAGE_KEYS.selectedPerson, found);
      renderPerson();
    }
  }

  function onPersonSelect() {
    state.selectedPerson = personSelectEl.value;
    personSearchEl.value = state.selectedPerson;
    saveString(STORAGE_KEYS.selectedPerson, state.selectedPerson);
    renderPerson();
  }

  function onTaskChecklistChange(evt) {
    var input = evt.target;
    if (!(input instanceof HTMLInputElement) || input.type !== "checkbox") return;
    var task = input.getAttribute("data-task");
    if (!task) return;

    state.selectedTasks[task] = input.checked;
    renderSummary();
    renderPerson();
  }

  function onExportIcs() {
    var built = buildIcsForSelectedPerson();
    if (!built) return;

    downloadBlob(built.blob, built.fileName);
    setStatus(buildExportStatusMessage(built, "Export ICS pret."), false);
  }

  function onShareIcs() {
    var built = buildIcsForSelectedPerson();
    if (!built) return;

    if (!canShareIcsFile()) {
      downloadBlob(built.blob, built.fileName);
      setStatus(buildExportStatusMessage(built, "Partage non disponible, fichier telecharge."), false);
      return;
    }

    navigator
      .share({
        title: built.fileName,
        text: "Calendrier pour " + state.selectedPerson,
        files: [built.file],
      })
      .then(function () {
        setStatus(buildExportStatusMessage(built, "Calendrier partage."), false);
      })
      .catch(function (err) {
        if (err && err.name === "AbortError") return;
        console.error(err);
        setStatus("Impossible de partager le calendrier.", true);
      });
  }

  function buildIcsForSelectedPerson() {
    var name = state.selectedPerson;
    if (!name) {
      setStatus("Selectionnez une personne avant l'export calendrier.", true);
      return null;
    }

    var personRecords = getVisiblePersonRecords(name);
    if (!personRecords.length) {
      setStatus("Aucun planning visible a exporter pour cette personne.", true);
      return null;
    }

    var groups = buildCalendarDayGroups(personRecords);
    if (!groups.days.length) {
      setStatus("Aucune date exploitable pour l'export calendrier.", true);
      return null;
    }

    var fileName = core.normalize(name).replace(/[^A-Z0-9]+/g, "-").replace(/^-+|-+$/g, "").toLowerCase() + "-planning.ics";
    var text = core.buildIcsText(name, groups.days);
    var blob = new Blob([text], { type: "text/calendar;charset=utf-8" });
    var file = typeof File === "function" ? new File([blob], fileName, { type: blob.type }) : null;

    return {
      fileName: fileName,
      blob: blob,
      file: file,
      eventCount: groups.days.length,
      skippedCount: groups.skippedCount,
    };
  }

  function buildCalendarDayGroups(personRecords) {
    return core.buildCalendarDayGroups(personRecords);
  }

  function canShareIcsFile() {
    if (typeof navigator === "undefined" || typeof navigator.share !== "function" || typeof File !== "function") {
      return false;
    }

    if (typeof navigator.canShare !== "function") {
      return false;
    }

    try {
      return navigator.canShare({
        files: [new File(["BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n"], "test.ics", { type: "text/calendar" })],
      });
    } catch (_) {
      return false;
    }
  }

  function buildExportStatusMessage(built, prefix) {
    var msg = prefix + " " + built.eventCount + " evenement(s).";
    if (built.skippedCount) {
      msg += " " + built.skippedCount + " jour(s) ignores faute de date exploitable.";
    }
    return msg;
  }

  function downloadBlob(blob, fileName) {
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 0);
  }


  function applyNameMap(rawName) {
    var key = nameKey(rawName);
    if (state.nameMap[key]) {
      return state.nameMap[key];
    }
    return rawName;
  }

  function updateAutoNameMap(records) {
    var changed = false;
    records.forEach(function (rec) {
      var source = toText(rec.rawPerson).trim();
      if (!source) return;
      var key = nameKey(source);
      if (!state.nameMap[key]) {
        state.nameMap[key] = source;
        changed = true;
      }
    });
    if (changed) {
      persistNameMap();
    }
  }

  function nameKey(name) {
    return normalize(name);
  }

  function persistNameMap() {
    try {
      localStorage.setItem(STORAGE_KEYS.nameMap, JSON.stringify(state.nameMap));
    } catch (_) {}
  }

  function loadNameMap() {
    var raw = loadString(STORAGE_KEYS.nameMap);
    if (!raw) return {};
    try {
      var parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") return {};

      var map = {};
      Object.keys(parsed).forEach(function (key) {
        var item = parsed[key];
        var target = "";
        if (typeof item === "string") {
          target = toText(item).trim();
        } else if (item && typeof item === "object") {
          target = toText(item.target).trim();
        }
        if (!target) return;
        map[key] = target;
      });
      return map;
    } catch (_) {
      return {};
    }
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

  function extractPeopleFromObservation(observation) {
    var text = toText(observation);
    if (!text) return [];

    var source = text;
    var colonIndex = source.indexOf(":");
    if (colonIndex >= 0) {
      source = source.slice(colonIndex + 1);
    }

    source = source.replace(/\bet\b/gi, ",");
    var candidates = source
      .split(/[;,/]/)
      .map(function (part) {
        return sanitizePerson(part);
      })
      .filter(Boolean);

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

  function deriveWeekMeta(weekLabel) {
    return {
      year: core.extractYearNumber(weekLabel),
      month: core.extractMonthNumber(weekLabel),
    };
  }

  function normalize(value) {
    return toText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function sortStrings(values) {
    return values.sort(function (a, b) {
      return a.localeCompare(b, "fr", { sensitivity: "base" });
    });
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

  function loadString(key) {
    try {
      return localStorage.getItem(key) || "";
    } catch (_) {
      return "";
    }
  }

  function saveString(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch (_) {}
  }

  function loadBool(key) {
    return loadString(key) === "1";
  }
})();
