'use strict';

var expect = require('chai').expect;
var parser = require('../lib/parser');
var xlsx = require('xlsx');

var workbook = xlsx.readFile("tests/data/testdata.xlsx");
var jsobject = parser.parseWorkbook(workbook, 0, 0);

describe("xlso", function(){
  describe("#parseWorkbook()", function(){
    it("parsed workbook schould be an array", function(){
      expect(jsobject).to.be.an('array');
    });
    it("parsed workbook schould have a length of 9", function(){
      expect(jsobject).to.have.length(9);
    });
    it("parsed workbook schould have a property 'Name'", function(){
      for (var i = 0; i < jsobject.length; i++) {
        expect(jsobject[i]).to.have.property('Name');
      }
    });
    it("parsed workbook schould have a property 'City'", function(){
      for (var i = 0; i < jsobject.length; i++) {
        expect(jsobject[i]).to.have.property('City');
      }
    });
    it("parsed workbook schould have a property 'Hobby'", function(){
      for (var i = 0; i < jsobject.length; i++) {
        expect(jsobject[i]).to.have.property('Hobby');
      }
    });
  })
});
