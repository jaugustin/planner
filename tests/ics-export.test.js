"use strict";

var test = require("node:test");
var assert = require("node:assert/strict");
var helpers = require("./helpers");

test("METRAL ICS export generates one event per day with person-task summaries", function () {
  var records = [];
  records = records.concat(helpers.mappedFixture("Planning JUIN envoyé.xlsx", 0));
  records = records.concat(helpers.mappedFixture("Planning Juillet envoyé.xlsx", 1));

  var personRecords = helpers.core.getVisiblePersonRecords(records, "METRAL");
  var groups = helpers.core.buildCalendarDayGroups(personRecords);
  var ics = helpers.core.buildIcsText("METRAL", groups.days);

  assert.ok(groups.days.length > 0, "METRAL should produce calendar days");
  assert.ok(ics.indexOf("SUMMARY:METRAL - ") >= 0, "ICS summaries should start with person-task title");
  assert.equal(ics.indexOf("SUMMARY:METRAL - Planning"), -1, "ICS should use task names when tasks exist");
});

test("cross-month July week exports July 1st correctly for METRAL", function () {
  var records = helpers.mappedFixture("Planning Juillet envoyé.xlsx", 0);
  var personRecords = helpers.core.getVisiblePersonRecords(records, "METRAL");
  var groups = helpers.core.buildCalendarDayGroups(personRecords);
  var ics = helpers.core.buildIcsText("METRAL", groups.days);

  assert.ok(ics.indexOf("DTSTART;VALUE=DATE:20260701") >= 0, "July workbook should export 2026-07-01");
  assert.equal(ics.indexOf("DTSTART;VALUE=DATE:20260501"), -1, "July workbook should not export 2026-05-01");
});
