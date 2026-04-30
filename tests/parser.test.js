"use strict";

var test = require("node:test");
var assert = require("node:assert/strict");
var helpers = require("./helpers");

var fixtureFiles = [
  "PLANNING MAGAZINE envoyé.xlsx",
  "PLANNING MARS envoyé.xlsx",
  "Planning MAGAZINE JUIN.xlsx",
  "Planning JUIN envoyé.xlsx",
  "Planning Magazine juillet envoyé.xlsx",
  "Planning Juillet envoyé.xlsx",
];

test("all fixture workbooks produce assignment records", function () {
  fixtureFiles.forEach(function (name, index) {
    var records = helpers.parseFixture(name, index);
    assert.ok(records.length > 0, name + " should produce records");
  });
});

test("METRAL is detected in the March and summer fixtures", function () {
  [
    "PLANNING MARS envoyé.xlsx",
    "Planning JUIN envoyé.xlsx",
    "Planning Juillet envoyé.xlsx",
  ].forEach(function (name, index) {
    var mapped = helpers.mappedFixture(name, index);
    var people = helpers.core.extractMappedPeople(mapped);
    assert.ok(people.indexOf("METRAL") >= 0, name + " should contain METRAL");
  });
});

test("obvious header text should not appear as a person", function () {
  var mapped = helpers.mappedFixture("Planning MAGAZINE JUIN.xlsx", 0);
  var people = helpers.core.extractMappedPeople(mapped);
    assert.equal(people.indexOf("PLANNING DES MAGAZINES - Du"), -1);
});
