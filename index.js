const path = require('path')
const ExcelJs = require('exceljs')

module.exports = class ExcelReader {
  /**
   * Creates an instance of ExcelReader.
   *
   * @param {string} filePath - Path to working excel file.
   * @param {string} [delimiter] - Used to parse cells holding string arrays.
   */
  constructor(filePath, delimiter = ';') {
    this._filePath = path.resolve(filePath)
    this._delimiter = delimiter
    this._workbook = undefined
    this._worksheetNames = undefined
  }

  _parseHeadings(row) {
    const headings = []
    row.eachCell({includeEmpty: true}, cell => {
      const heading = cell.value
      const isArray = /\[\]$/.test(heading)
      headings.push({
        key: !isArray ? heading : heading.slice(0, -2),
        isArray,
      })
    })
    return headings
  }

  _parseRecord(headings, row) {
    return headings.reduce((record, columnData, colCount) => {
      const {key, isArray} = columnData
      const rawValue = row.getCell(colCount + 1).value
      record[key] = !isArray
        ? rawValue
        : !rawValue
        ? []
        : rawValue.split(this._delimiter).trim()
      return record
    }, {})
  }

  /**
   * Get worksheet names.
   *
   * @returns {Array} Names of every worksheet.
   */
  async getWorksheetNames() {
    try {
      if (!this._workbook) {
        this._workbook = new ExcelJs.Workbook()
        await this._workbook.xlsx.readFile(this._filePath)
      }
      this._worksheetNames = []
      this._workbook.eachSheet(worksheet => {
        this._worksheetNames.push(worksheet.name)
      })
      return this._worksheetNames
    } catch (error) {
      throw error
    }
  }

  /**
   * Get sets of data records from each worksheet within the work Excel file.
   *
   * The async Array.reduce() function is made possible by using the following guide: https://css-tricks.com/why-using-reduce-to-sequentially-resolve-promises-works/.
   *
   * @returns {Object} Javascript object holding records from each worksheet, with the worksheet names as keys.
   */
  async getWorkbook() {
    try {
      await this.getWorksheetNames()
      const recordset = {}
      await this._worksheetNames.reduce((asyncAccumulator, worksheetName) => {
        return asyncAccumulator
          .then(() => this.getWorksheet(worksheetName))
          .then(records => {
            recordset[worksheetName] = records
            return Promise.resolve()
          })
          .catch(error => Promise.reject(error))
      }, Promise.resolve())
      return recordset
    } catch (error) {
      throw error
    }
  }

  /**
   * Get data records from a particular worksheet.
   *
   * @param {string} worksheetName - Name of worksheet to extract data from.
   * @returns {Array} Array of javascript objects holding an object (data record) with column heading as keys.
   */
  async getWorksheet(worksheetName) {
    try {
      if (!this._workbook) {
        this._workbook = new ExcelJs.Workbook()
        await this._workbook.xlsx.readFile(this._filePath)
      }
      const worksheet = this._workbook.getWorksheet(worksheetName)
      const headings = this._parseHeadings(worksheet.getRow(1))
      const records = []
      for (let rowCount = 2; rowCount <= worksheet.actualRowCount; rowCount++) {
        records.push(this._parseRecord(headings, worksheet.getRow(rowCount)))
      }
      return records
    } catch (error) {
      throw error
    }
  }
}
