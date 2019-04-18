# Excel Reader (excel-reader)

Simple node.js excel file (.xlsx) data extractor for
"**row oriented**" data files with known number of sheets

## Installation

```bash
npm install --save excel-reader
```

## Usage

```javascript
// SETUP
const ExcelReader = require('excel-reader')
const filePath = './directory/filename.xlsx'
/*
 * The 'sheetNames' array contains values to be used when calling
 * .getWorkSheet(sheetName).  These values are not actually
 * checked against the actual excel worksheet names
 *
 * The ordering is important though, in the sense that when
 * .getWorkSheet('oranges') is called, the desired data is actually
 * the data from the second sheet of the excel file.  Although the
 * actual sheet name may actually be 'smelly socks'.
 */
const sheetNames = ['apples', 'oranges', 'watermelons']
const excelReader = new ExcelReader(filePath, sheetNames)

// DATA EXTRACTION
// get the full workbook data in an object
// with sheet names as object keys
excelReader
  .getWorkBook()
  .then(datasets => {
    console.log(datasets)
    // get the data in array from from the 2nd
    // sheet of the workbook
    return excelReader.getWorkSheet('oranges')
  })
  .then(dataset => {
    console.log(dataset)
    return Promise.resolve()
  })
  .catch(error => {
    console.log(error)
  })
```

## Notes

Uses [excel-as-json](https://www.npmjs.com/package/excel-as-json) to
parse data, so check out its docs about the available source data file
structure.
