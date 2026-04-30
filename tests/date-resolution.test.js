"use strict";

var test = require("node:test");
var assert = require("node:assert/strict");
var core = require("../planning-core");

test("detects all French months explicitly", function () {
  assert.equal(core.extractMonthNumber("JANVIER 2026"), 1);
  assert.equal(core.extractMonthNumber("FEVRIER 2026"), 2);
  assert.equal(core.extractMonthNumber("MARS 2026"), 3);
  assert.equal(core.extractMonthNumber("AVRIL 2026"), 4);
  assert.equal(core.extractMonthNumber("MAI 2026"), 5);
  assert.equal(core.extractMonthNumber("JUIN 2026"), 6);
  assert.equal(core.extractMonthNumber("JUILLET 2026"), 7);
  assert.equal(core.extractMonthNumber("AOUT 2026"), 8);
  assert.equal(core.extractMonthNumber("SEPTEMBRE 2026"), 9);
  assert.equal(core.extractMonthNumber("OCTOBRE 2026"), 10);
  assert.equal(core.extractMonthNumber("NOVEMBRE 2026"), 11);
  assert.equal(core.extractMonthNumber("DECEMBRE 2026"), 12);
});

test("does not infer MAI from unrelated words like SEMAINE or DIMANCHE", function () {
  assert.equal(core.extractMonthNumber("SEMAINE 4"), 0);
  assert.equal(core.extractMonthNumber("DIMANCHE 5"), 0);
});

test("resolves cross-month week labels correctly for July 1st", function () {
  var date = core.resolveRecordDate({
    dayNumber: 1,
    dayLabel: "MERCREDI 1er",
    dateContextLabel: "SEMAINE 4 - Du lundi 29 JUIN au dimanche 5 JUILLET 2026",
    weekLabel: "SEMAINE 4 - Du lundi 29 JUIN au dimanche 5 JUILLET 2026",
    sheetName: "SEMAINE 4",
    fileName: "Planning Juillet envoyé.xlsx",
    sortMonth: 7,
    sortYear: 2026,
  });

  assert.equal(date && date.toISOString().slice(0, 10), "2026-07-01");
});

test("resolves explicit day/month labels correctly", function () {
  var date = core.resolveRecordDate({
    dayNumber: 29,
    dayLabel: "LUNDI 29/06",
    dateContextLabel: "SEMAINE 4 - Du lundi 29 JUIN au dimanche 5 JUILLET 2026",
    weekLabel: "SEMAINE 4 - Du lundi 29 JUIN au dimanche 5 JUILLET 2026",
    sheetName: "SEMAINE 4",
    fileName: "Planning Juillet envoyé.xlsx",
    sortMonth: 7,
    sortYear: 2026,
  });

  assert.equal(date && date.toISOString().slice(0, 10), "2026-06-29");
});
