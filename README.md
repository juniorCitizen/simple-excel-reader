# Simple Excel Reader (simple-excel-reader)

Simple node.js excel file (.xlsx) data extractor for
"**row oriented**" data files

## Installation

```bash
npm install --save simple-excel-reader
```

## Usage

```javascript
// SETUP
const ExcelReader = require('simple-excel-reader')
const filePath = './directory/filename.xlsx'
const excelReader = new ExcelReader(filePath)

// DATA EXTRACTION
// get the full workbook recordsets
excelReader
  .getWorkbook()
  .then(recordsets => {
    console.log(recordsets)
    // get the name of each worksheet
    return excelReader.getWorksheetNames()
  })
  .then(worksheetNames => {
    console.log(worksheetNames)
    // get the records from a particular worksheet
    return excelReader.getWorksheet(worksheetNames[2])
  })
  .then(records => console.log(records))
  .catch(error => console.error(error))
```

## API

### **new ExcelReader(filePath, delimiter)**

### **instance.getWorkbook()**

### **instance.getWorksheetNames()**

### **instance.getWorksheet(worksheetName)**
