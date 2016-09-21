'use strict';

// function to get the value of the current cell
var getCurrentCell = function(col, row, i, j){
  //get column name e.g. A, T etc. based on column position
  var c = String.fromCharCode(col.charCodeAt(0) + j);
  //get row/line number e.g. 1 for the first line etc. based on row position
  var r = i + 1;
  //get the cell name by concaternate column name and row number;
  return String(c + r);
};

/**
* workbook the workbook that contains the sheet to map
* sline the position of the sheet of the workbook to map
* hline the position of the header line.
*/
var parseWorkbook = function(workbook, sline, hline){
  //get the sheetname at position 'sline' (start at 0)
  var sheetname = workbook.SheetNames[sline];
  //get the the worksheet contains data to manipulate
  var worksheet = workbook.Sheets[sheetname];
  //get the span of the worksheet. e.g. A1:X14
  var span = worksheet["!ref"];
  //split the span to get the extrem values
  var cells = span.split(":");
  //get column range as array
  var cols = [cells[0].match(/[A-Z]+/g)[0], cells[1].match(/[A-Z]+/g)[0]];
  // get row range as array
  var rows = [cells[0].match(/[0-9]+/g)[0], cells[1].match(/[0-9]+/g)[0]];
  // check if the specified header is out of bound
  if(hline + 1 < rows[0] || hline >= rows[1]){
    // throw an exception the signalize that call set a wrong value
    throw new Error("the specified header is out of bound");
  }
  //Number of rows of the excel file
  var rlength = rows[1] - rows[0] + 1;
  //Number of columns of the excel file
  var clength = cols[1].charCodeAt(0) - cols[0].charCodeAt(0) + 1;
  var header = [];
  var toObject = function(arr) {
    var obj = {};
    var key = "";
    for (var i = 0; i < arr.length; i++){
      key = header[i].trim();
      var cRow = arr[i];
      obj[key] = cRow;
    }
    return obj;
  };
  var row;
  var result = [];
  for(var i = hline; i < rlength; i++){
    row = [];
    for(var j = 0; j < clength; j++){
      var cell = getCurrentCell(cols[0], row, i, j);
      var cellValue = (worksheet[cell]) ? worksheet[cell].v : "";
      if(i !== hline){
        row.push(cellValue);
      }else{
        //add the header row
        header.push(cellValue.toString());
      }
    }
    if(i !== hline){
      result.push(toObject(row));
    }
  }
  return result;
};

var parser = { parseWorkbook : parseWorkbook };

module.exports = parser;
