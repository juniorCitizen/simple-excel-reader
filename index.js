const path = require('path')
const excelAsJson = require('excel-as-json').processFile

module.exports = class ExcelReader {
  constructor(filePath, sheetNames) {
    this._filePath = path.resolve(filePath)
    this._sheetNames = sheetNames
    this._refKeys = this._sheetNames.reduce((_refKeys, sheetName, index) => {
      _refKeys[sheetName] = index + 1
      return _refKeys
    }, {})
  }

  async getWorkBook() {
    try {
      const mapFn = sheetName => {
        const sheetIndex = this._refKeys[sheetName].toString()
        return this._parseWorkSheet(sheetIndex)
      }
      const datasets = await Promise.all(this._sheetNames.map(mapFn))
      return this._sheetNames.reduce((_workbook, sheetName, index) => {
        _workbook[sheetName] = datasets[index]
        return _workbook
      }, {})
    } catch (error) {
      throw error
    }
  }

  async getWorkSheet(sheetName) {
    try {
      if (this._sheetNames.includes(sheetName)) {
        const sheetIndex = this._refKeys[sheetName].toString()
        return await this._parseWorkSheet(sheetIndex)
      } else {
        throw new Error(`unknown sheet name: ${sheetName}`)
      }
    } catch (error) {
      throw error
    }
  }

  _parseWorkSheet(sheetIndex) {
    return new Promise((resolve, reject) => {
      const opt = {sheet: sheetIndex}
      excelAsJson(this._filePath, undefined, opt, (error, data) => {
        if (error) reject(error)
        else resolve(data)
      })
    })
  }
}
