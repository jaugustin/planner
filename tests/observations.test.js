"use strict";

var test = require("node:test");
var assert = require("node:assert/strict");
var helpers = require("./helpers");

test("METRAL gets the shared March formation observation", function () {
  var mapped = helpers.mappedFixture("PLANNING MARS envoyé.xlsx", 0);
  var metralRecords = mapped.filter(function (rec) {
    return rec.person === "METRAL" && rec.observation;
  });

  var observations = metralRecords.map(function (rec) {
    return rec.observation;
  });

  assert.ok(
    observations.indexOf("Formation 11,12/03 : Giguere, Metral, Boï, Tibet, Gosse") >= 0,
    "METRAL should carry the shared formation observation"
  );
});
