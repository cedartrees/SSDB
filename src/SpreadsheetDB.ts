/**
 * SpreadsheetDB クラスは、Google スプレッドシートをデータベースのように操作するためのクラスです。
 * @class SpreadsheetDB
 */
class SpreadsheetDB {

  // internal properties
  spreadsheetId: string;
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheetMap: {[key: string]: SheetTable};

  /**
   * SpreadsheetDB クラスのインスタンスを作成します。
   * @param {string} spreadsheetId - スプレッドシートのID。指定しない場合、アクティブなスプレッドシートのIDが使用されます。
   */
  constructor(spreadsheetId: string) {
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
  insert(sheetName: string, rowValues: {[key: string]: any}): number {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.insert(rowValues);
  }

  /**
   * 指定したシートに複数の行を追加します。
   * @param {string} sheetName - 行を追加するシートの名前。
   * @param {Array<Object>} rowValuesArray - 追加する行のデータの配列。各要素は、キーが列名、値がセルの値のオブジェクトです。
   * @returns {number} 追加された行数。
   */
  insertAll(sheetName: string, rowValuesArray: Array<object>): number {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.insertAll(rowValuesArray);
  }

  /**
   * 指定したシートから、プライマリキーに基づいて行を検索します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。入力しない場合はソートされません。
   * @returns {Array<Object>} 検索に一致した行のデータ。一致する行がない場合は空の配列。
   */
  selectByPk(sheetName: string, pkColumnName: string, pkValue: any, sortBy: { column: string; order: "ASC" | "DESC"; }): Array<object> {
    const sheetTable = this.getSheetTable_(sheetName);

    if (sortBy) {
      return sheetTable.selectByPkSorted(pkColumnName, pkValue, sortBy);
    }

    return sheetTable.selectByPk(pkColumnName, pkValue);
  }

  /**
   * 指定したシートから、指定した列の値に基づいて行を検索します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {string} columnName - 検索に使用する列名。
   * @param {*} value - 検索する列の値。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。入力しない場合はソートされません。
   * @returns {Array<Object>} 検索に一致した行のデータの配列。一致する行がない場合は空の配列。
   */
  selectByColumn(sheetName: string, columnName: string, value: any, sortBy: { column: string; order: "ASC" | "DESC"; }): Array<object> {
    const sheetTable = this.getSheetTable_(sheetName);

    if (sortBy) {
      return sheetTable.selectByColumnSorted(columnName, value, sortBy);
    }
    return sheetTable.selectByColumn(columnName, value);
  }

  /**
   * 指定したシートから、指定した列の値に基づいて行を検索し、結果を指定した条件でソートします。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {string} columnName - 検索に使用する列名。
   * @param {*} value - 検索する列の値。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。
   * @returns {Array<Object>} 検索に一致した行のデータの配列。一致する行がない場合は空の配列。
   * @deprecated selectByColumn() に sortBy を指定することで代替できます。
   */
  selectByColumnSorted(sheetName: string, columnName: string, value: any, sortBy: { column: string; order: "ASC" | "DESC"; }): Array<object> {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectByColumnSorted(columnName, value, sortBy);
  }

  /**
   * 指定したシートから、指定した複数の列の値に基づいて行を検索します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。入力しない場合はソートされません。
   * @returns {Array<Object>} 検索に一致した行のデータの配列。一致する行がない場合は空の配列。
   */
  selectByColumns(sheetName: string, criteria: object, sortBy: { column: string; order: "ASC" | "DESC"; }): Array<object> {
    const sheetTable = this.getSheetTable_(sheetName);

    if (sortBy) {
      return sheetTable.selectByColumnsSorted(criteria, sortBy);
    }
    return sheetTable.selectByColumns(criteria);
  }

  /**
   * 指定したシートから、全てのデータを取得します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。入力しない場合はソートされません。
   * @returns {Array<Object>} シートの全行のデータの配列。一致する行がない場合は空の配列。
   */
  selectAll(sheetName: string, sortBy: { column: string; order: "ASC" | "DESC"; }): Array<object> {
    const sheetTable = this.getSheetTable_(sheetName);

    if (sortBy) {
      return sheetTable.selectAllSorted(sortBy);
    }
    return sheetTable.selectAll();
  }

  /**
   * 指定したシートから、指定した列の最大値を取得します。
   * @param {string} columnName - 最大値を取得する列名。
   * @returns {number|null} 指定した列の最大値。列が存在しない場合は null。
   */
  selectMax(sheetName, columnName: string): number | null {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectMax(columnName);
  }

  /**
   * 指定したシートから、指定した列の値をインクリメントして更新、取得します。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {string} columnName - インクリメントする列名。
   * @param {number} increment - インクリメントする値。
   * @returns {number|null} インクリメント後の値。列が存在しない場合は null。
   */
  selectByPkAndIncrement(sheetName, pkColumnName: string, pkValue: any, columnName: string, increment: number): number | null {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.selectByPkAndIncrement(pkColumnName, pkValue, columnName, increment);
  }

  /**
   * 指定したシートの指定したプライマリキーに基づいて行を更新します。
   * @param {string} sheetName - 行を更新するシートの名前。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {Object} rowValues - 更新する行のデータ。キーは列名、値はセルの値です。
   * @returns {Array<Object>} 更新された行のデータ。一致する行がない場合は空の配列。
   */
  updateByPk(sheetName: string, pkColumnName: string, pkValue: any, rowValues: object): Array<object> {
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
   * @returns {Array<Object>} 更新された行のデータ。一致する行がない場合は空の配列。
   */
  updateItemByPk(sheetName: string, pkColumnName: string, pkValue: any, columnName: string, value: any): Array<object> {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.updateItemByPk(pkColumnName, pkValue, columnName, value);
  }

  /**
   * 指定したシートの指定した列の値に基づいて行を更新します。
   * @param {string} sheetName - 行を更新するシートの名前。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {String} columnName - 更新する列の名前。
   * @param {*} value - 更新する列の値。
   * @returns {Array<Object>} 更新された行のデータ。一致する行がない場合は空の配列。
   */
  updateItemByColumns(sheetName: string, criteria: object, columnName: string, value: any): Array<object> {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.updateItemByColumns(criteria, columnName, value);
  }

  /**
   * 指定したシートの指定した列の値に基づいて行を更新します。
   * @param {string} sheetName - 行を更新するシートの名前。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {Object} columnValues - 更新対象となるカラムと値を持つオブジェクト。
   * @returns {Array<Object>} 更新された行のデータ。一致する行がない場合は空の配列。
   */
  updateItemsByColumns(sheetName: string, criteria: object, columnValues: Object): Array<object> {
    const sheetTable = this.getSheetTable_(sheetName);
    return sheetTable.updateItemsByColumns(criteria, columnValues);
  }

  // private methods

  /**
   * 指定したシート名の SheetTable オブジェクトを取得します。
   * @param {string} sheetName - 取得する SheetTable オブジェクトのシート名。
   * @returns {SheetTable} 指定したシート名の SheetTable オブジェクト。
   */
  getSheetTable_(sheetName: string): SheetTable {
    if (this.sheetMap[sheetName]) {
      return this.sheetMap[sheetName];
    }
    const sheet = new SheetTable(this.spreadsheet, sheetName);
    this.sheetMap[sheetName] = sheet;
    return sheet;
  }
}