"use strict";

var path = require("path");
var XLSX = require("xlsx");
var core = require("../planning-core");

function fixturePath(name) {
  return path.join(__dirname, "..", "example files", name);
}

function loadWorkbook(name) {
  return XLSX.readFile(fixturePath(name), { cellDates: false });
}

function parseFixture(name, fileOrder) {
  var workbook = loadWorkbook(name);
  return core.parseWorkbook(workbook, name, typeof fileOrder === "number" ? fileOrder : 0);
}

function mappedFixture(name, fileOrder, nameMap) {
  return core.mapRecords(parseFixture(name, fileOrder), nameMap || {});
}

module.exports = {
  core: core,
  fixturePath: fixturePath,
  loadWorkbook: loadWorkbook,
  parseFixture: parseFixture,
  mappedFixture: mappedFixture,
};
