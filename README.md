[![Known Vulnerabilities](https://snyk.io/test/npm/xlso/badge.svg)](https://snyk.io/test/npm/xlso) [![bitHound Overall Score](https://www.bithound.io/github/nanachimi/xlso/badges/score.svg)](https://www.bithound.io/github/nanachimi/xlso) [![bitHound Code](https://www.bithound.io/github/nanachimi/xlso/badges/code.svg)](https://www.bithound.io/github/nanachimi/xlso) [![bitHound Dependencies](https://www.bithound.io/github/nanachimi/xlso/badges/dependencies.svg)](https://www.bithound.io/github/nanachimi/xlso/master/dependencies/npm)

# xlso
Node.js package for parsing excel workbooks into an Javascript array of rows. Excel workbooks based on the node.js package [xlsx](https://github.com/SheetJS/js-xlsx).

## how to install the xlso?

```Shell
$ npm install xlso --save
```

## How to use the package?

Let suppose the following table is saved as `testdata.xlsx`.

| Id    | Name  | City  |
| ----- | ----- | ----- |
|  1    | Micky | Loum  |
|  2    | Lapiro| Mbanga|

The following snippet show how to parse `testdata.xlsx` in order to get all data.

```Javascript
var xlso = require('xlso');
var xlsx = require('xlsx');
// the xlsx's workbook to parse by the xlso package.
var workbook = xlsx.readFile("testdata.xlsx");
// parse the workbook to a js array
var rows = xlso.parseWorkbook(workbook, 0, 0);
// display your js array
rows.forEach(function(item, index){
  console.log("Id: " + item.Id + ", Name: " + item.Name + ", City: " + item.City);
});
```

## API Specifications
`xlso` expose one API that expect three parameters:
* the `first parameter` is the workbook that was generated by the package xlsx (see section below).
* the `second parameter` is the position of the sheet in the given workbook (the index 0 is the first sheet).
* the third and `last parameter` is the position of the header row (position 0 means also first row).

PS: the header row must be the first row.

## How to run unit-tests?

```Shell
$ npm test
```
