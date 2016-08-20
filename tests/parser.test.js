'use strict';

var expect = require('chai').expect;
var describe = require('mocha').describe;
var it = require('mocha').it;
var parser = require('../lib/parser');
var xlsx = require('xlsx');

var workbook = xlsx.readFile("tests/data/testdata.xlsx");
var jsobject = parser.parseWorkbook(workbook, 0, 0);

describe("xlso", function(){
  describe("#parseWorkbook()", function(){
    it("parsed workbook schould be an array", function(){
      expect(jsobject).to.be.an('array');
    });
    it("parsed workbook schould have a length of 5", function(){
      expect(jsobject).to.have.length(5);
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
    it("parsed skills schould have a valid 'Id' value", function(){
      expect(jsobject[0].Id).to.equals(1);
      expect(jsobject[1].Id).to.equals(2);
      expect(jsobject[2].Id).to.equals(3);
      expect(jsobject[3].Id).to.equals(4);
      expect(jsobject[4].Id).to.equals(5);
    });
    it("parsed skills schould have a valid 'Name' value", function(){
      expect(jsobject[0].Name).to.equals('Micky');
      expect(jsobject[1].Name).to.equals('Tonton');
      expect(jsobject[2].Name).to.equals('Kobe');
      expect(jsobject[3].Name).to.equals('Lapiro');
      expect(jsobject[4].Name).to.equals('Bougaps');
    });
    it("parsed skills schould have a valid 'City' value", function(){
      expect(jsobject[0].City).to.equals('Loum');
      expect(jsobject[1].City).to.equals('Akonolinga');
      expect(jsobject[2].City).to.equals('Douala');
      expect(jsobject[3].City).to.equals('Mbanga');
      expect(jsobject[4].City).to.equals('Yaoundé');
    });
    it("parsed skills schould have a valid 'Hobby' value", function(){
      expect(jsobject[0].Hobby).to.equals('Basketball');
      expect(jsobject[1].Hobby).to.equals('Football');
      expect(jsobject[2].Hobby).to.equals('Tennis');
      expect(jsobject[3].Hobby).to.equals('Jeux de Dame');
      expect(jsobject[4].Hobby).to.equals('Jeux de Carte');
    });
  });
});
