/**
 * SpreadsheetDB クラスは、Google スプレッドシートをデータベースのように操作するためのクラスです。
 * @class SpreadsheetDB
 */
class SpreadsheetDB {
  /**
   * SpreadsheetDB クラスのインスタンスを作成します。
   * @param {string} spreadsheetId - スプレッドシートのID。指定しない場合、アクティブなスプレッドシートのIDが使用されます。
   */
  constructor(spreadsheetId) {
    if (!spreadsheetId) {
      this.spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    } else {
      this.spreadsheetId = spreadsheetId;
    }
    try {
      this.spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
    } catch (e) {
      console.log(e);
      throw new Error("スプレッドシートが見つかりません");
    }

    this.sheetMap = {};
  }

  /**
   * 指定したシートに行を追加します。
   * @param {string} sheetName - 行を追加するシートの名前。
   * @param {Object} rowValues - 追加する行のデータ。キーは列名、値はセルの値です。
   * @returns {number} 追加された行数。
   */
  insert(sheetName, rowValues) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.insert(rowValues);
  }

  /**
   * 指定したシートに複数の行を追加します。
   * @param {string} sheetName - 行を追加するシートの名前。
   * @param {Array<Object>} rowValuesArray - 追加する行のデータの配列。各要素は、キーが列名、値がセルの値のオブジェクトです。
   * @returns {number} 追加された行数。
   */
  insertAll(sheetName, rowValuesArray) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.insertAll(rowValuesArray);
  }

  /**
   * 指定したシートから、プライマリキーに基づいて行を検索します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @returns {Array<Object>|null} 検索に一致した行のデータ。一致する行がない場合は null。
   */
  selectByPk(sheetName, pkColumnName, pkValue) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectByPk(pkColumnName, pkValue);
  }

  /**
   * 指定したシートから、指定した列の値に基づいて行を検索します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {string} columnName - 検索に使用する列名。
   * @param {*} value - 検索する列の値。
   * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
   */
  selectByColumn(sheetName, columnName, value) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectByColumn(columnName, value);
  }

  /**
   * 指定したシートから、指定した列の値に基づいて行を検索し、結果を指定した条件でソートします。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {string} columnName - 検索に使用する列名。
   * @param {*} value - 検索する列の値。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。
   * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
   */
  selectByColumnSorted(sheetName, columnName, value, sortBy) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectByColumnSorted(columnName, value, sortBy);
  }

  /**
   * 指定したシートから、指定した複数の列の値に基づいて行を検索します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
   */
  selectByColumns(sheetName, criteria) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectByColumns(criteria);
  }

  /**
   * シートから、全てのデータを取得します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @returns {Array<Object>|null} シートの全行のデータの配列。
   */
  selectAll(sheetName) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectAll();
  }

  selectMax(sheetName, columnName) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectMax(columnName);
  }

  selectByPkAndIncrement(sheetName, pkColumnName, pkValue, columnName, increment) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectByPkAndIncrement(pkColumnName, pkValue, columnName, increment);
  }

  /**
   * 指定したシートの指定したプライマリキーに基づいて行を更新します。
   * @param {string} sheetName - 行を更新するシートの名前。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {Object} rowValues - 更新する行のデータ。キーは列名、値はセルの値です。
   * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
   */
  updateByPk(sheetName, pkColumnName, pkValue, rowValues) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.updateByPk(pkColumnName, pkValue, rowValues);
  }

  /**
   * 指定したシートの指定したプライマリキーに基づいて、特定の列の値を更新します。
   * @param {string} sheetName - 行を更新するシートの名前。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {string} columnName - 更新する列の名前。
   * @param {*} value - 更新する列の値。
   * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
   */
  updateItemByPk(sheetName, pkColumnName, pkValue, columnName, value) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.updateItemByPk(pkColumnName, pkValue, columnName, value);
  }

  /**
   * 指定したシートの指定した列の値に基づいて行を更新します。
   * @param {string} sheetName - 行を更新するシートの名前。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {String} columnName - 更新する列の名前。
   * @param {*} value - 更新する列の値。
   * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
   */
  updateItemByColumns(sheetName, criteria, columnName, value) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.updateItemByColumns(criteria, columnName, value);
  }

  /**
   * 指定したシートの指定した列の値に基づいて行を更新します。
   * @param {string} sheetName - 行を更新するシートの名前。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {Objext} columnValues - 更新対象となるカラムと値を持つオブジェクト。
   * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
   */
  updateItemsByColumns(sheetName, criteria, columnValues) {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.updateItemsByColumns(criteria, columnValues);
  }

  // private methods
  getSheetTable_(sheetName) {
    if (this.sheetMap[sheetName]) {
      return this.sheetMap[sheetName];
    }
    const sheet = new SheetTable(this.spreadsheet, sheetName);
    this.sheetMap[sheetName] = sheet;
    return sheet;
  }
}
