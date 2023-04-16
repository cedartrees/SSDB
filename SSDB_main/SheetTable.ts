/**
 * SheetTable クラスは、Google スプレッドシートのシートを操作するためのクラスです。
 */
class SheetTable {

  // internal properties
  private sheetName: string;
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private headers: string[];
  private columnIndexMap: {[key: string]: number};
  private ROW_INDEX_CONVERT_NUM: number = 2;

  /**
   * SheetTable クラスのインスタンスを作成します。
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - Google スプレッドシートのインスタンス。
   * @param {string} sheetName - シートの名前。
   */
  constructor(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: string) {
    this.sheetName = sheetName;
    try {
      const sheet = spreadsheet.getSheetByName(this.sheetName);
      if (!sheet) {
        throw new Error("シートが見つかりません");
      }
      this.sheet = sheet;
    } catch (e) {
      console.log(e);
      throw new Error("シートが見つかりません");
    }
    this.headers = this.sheet
      .getRange(1, 1, 1, this.sheet.getLastColumn())
      .getValues()[0];
    this.columnIndexMap = this.createColumnIndexMap(this.headers);
  }

  /**
   * シートに行を追加します。
   * @param {Object} rowValues - 追加する行のデータ。キーは列名、値はセルの値です。
   * @returns {number} 追加された行数。
   */
  insert(rowValues: object): number {
    const record = this.createRecord(rowValues, this.getColumnIndexMap());
    this.getSheet().appendRow(record);
    return 1;
  }

  /**
   * シートに複数の行を追加します。
   * @param {Array<Object>} rowValuesArray - 追加する行のデータの配列。各要素は、キーが列名、値がセルの値のオブジェクトです。
   * @returns {number} 追加された行数。
   */
  insertAll(rowValuesArray: Array<object>): number {
    const records = rowValuesArray.map((rowValues) =>
      this.createRecord(rowValues, this.getColumnIndexMap())
    );
    this.getSheet()
      .getRange(
        this.getSheet().getLastRow() + this.ROW_INDEX_CONVERT_NUM,
        1,
        records.length,
        records[0].length
      )
      .setValues(records);
    return records.length;
  }

  /**
   * シートから、プライマリキーに基づいて行を検索します。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @returns {Array<Object>} 検索に一致した行のデータ。一致する行がない場合は空の配列。
   */
  selectByPk(pkColumnName: string, pkValue: any): Array<object> {
    const index = this.getIndexByColumnName(pkColumnName);
    const data = this.getDataset();
    const record = data.find((row) => this.valuesMatch(row[index], pkValue));
    if (!record) return [];

    return this.convertRecordToObj([record]);
  }

  /**
   * シートから、プライマリキーに基づいて行を検索し、結果を指定した条件でソートします。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。
   * @returns {Array<Object>} 検索に一致した行のデータの配列。一致する行がない場合は空の配列。
   */
  selectByPkSorted(pkColumnName: string, pkValue: any, sortBy: { column: string; order: "ASC" | "DESC"; }): Array<object> {
    const records = this.selectByPk(pkColumnName, pkValue);
    if (records.length === 0) return [];

    const sortedRecords = this.sortData(
      records,
      sortBy);
    return this.convertRecordToObj(sortedRecords);
  }

  /**
   * シートから、指定した列の値に基づいて行を検索します。
   * @param {string} columnName - 検索に使用する列名。
   * @param {*} value - 検索する列の値。
   * @returns {Array<Object>} 検索に一致した行のデータの配列。一致する行がない場合は空の配列。
   */
  selectByColumn(columnName: string, value: any): Array<object> {
    const index = this.getIndexByColumnName(columnName);
    const data = this.getDataset();
    const records = data.filter((row) => {
      return this.valuesMatch(row[index], value);
    });

    if (records.length === 0) return [];

    return this.convertRecordToObj(records);
  }

  /**
   * シートから、指定した列の値に基づいて行を検索し、結果を指定した条件でソートします。
   * @param {string} columnName - 検索に使用する列名。
   * @param {*} value - 検索する列の値。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。
   * @returns {Array<Object>} 検索に一致した行のデータの配列。一致する行がない場合は null。
   */
  selectByColumnSorted(columnName: string, value: any, sortBy: { column: string; order: "ASC" | "DESC"; }): Array<object> {
    const records = this.selectByColumn(columnName, value);
    if (records.length === 0) return [];

    const sortedRecords = this.sortData(
      records,
      sortBy);
    return this.convertRecordToObj(sortedRecords);
  }

  /**
   * シートから、指定した複数の列の値に基づいて行を検索します。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @returns {Array<Object>} 検索に一致した行のデータの配列。一致する行がない場合は空の配列。
   */
  selectByColumns(criteria: object): Array<object> {
    const data = this.getDataset();
    const records = data.filter(
      (row) =>
        Object.keys(criteria).length === 0 ||
        Object.entries(criteria).every(([columnName, value]) => {
          const index = this.getIndexByColumnName(columnName);
          return this.valuesMatch(row[index], value);
        })
    );

    if (records.length === 0) return [];

    return this.convertRecordToObj(records);
  }

  /**
   * シートから、指定した複数の列の値に基づいて行を検索し、結果を指定した条件でソートします。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。
   * @returns {Array<Object>} 検索に一致した行のデータの配列。一致する行がない場合は空の配列。
   */
  selectByColumnsSorted(criteria: object, sortBy: { column: string; order: "ASC" | "DESC"; }): Array<object> {
    const records = this.selectByColumns(criteria);
    if (records.length === 0) return [];

    const sortedRecords = this.sortData(
      records,
      sortBy);
    return this.convertRecordToObj(sortedRecords);
  }

  /**
   * シートから、全てのデータを取得します。
   * @return {Array<Object>} シートの全行のデータの配列。一致する行がない場合は空の配列。
   */
  selectAll(): Array<object> {
    const data = this.getDataset();
    return this.convertRecordToObj(data);
  }

  selectAllSorted(sortBy) {
    const records = this.selectAll();
    if (records.length === 0) return [];

    const sortedRecords = this.sortData(
      records,
      sortBy);
    return this.convertRecordToObj(sortedRecords);
  }

  /**
   * シートから、指定した列の最大値を取得します。
   * @param {string} columnName - 最大値を取得する列名。
   * @returns {number|null} 指定した列の最大値。列が存在しない場合は null。
   */
  selectMax(columnName: string): number | null {
    if (typeof columnName !== "string" || columnName === "") {
      throw new Error("Invalid columnName");
    }

    const index = this.getIndexByColumnName(columnName);
    if (index === undefined) {
      throw new Error(`Column "${columnName}" not found`);
    }

    const data = this.getDataset();
    if (data.length === 0) return null;

    const initialMax = data[0][index];
    const max = data.reduce((currentMax, row) => {
      const value = row[index];
      if (typeof value === "number") {
        return Math.max(currentMax, value);
      }
      return currentMax;
    }, initialMax);

    return max;
  }

  /**
   * シートから、指定した列の値をインクリメントして更新、取得します。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {string} columnName - インクリメントする列名。
   * @param {number} increment - インクリメントする値。
   * @returns {number|null} インクリメント後の値。列が存在しない場合は null。
   */
  selectByPkAndIncrement(pkColumnName: string, pkValue: any, columnName: string, increment: number = 1): number | null {
    if (typeof pkColumnName !== "string" || pkColumnName === "") {
      throw new Error("Invalid pkColumnName");
    }

    if (typeof columnName !== "string" || columnName === "") {
      throw new Error("Invalid columnName");
    }

    if (typeof increment !== "number" || isNaN(increment)) {
      throw new Error("Invalid increment value");
    }

    const pkColumnIndex = this.getIndexByColumnName(pkColumnName);
    if (pkColumnIndex === undefined) {
      throw new Error(`Column "${pkColumnName}" not found`);
    }

    const columnIndex = this.getIndexByColumnName(columnName);
    if (columnIndex === undefined) {
      throw new Error(`Column "${columnName}" not found`);
    }

    const data = this.getDataset();

    const recordIndex = data.findIndex((row) =>
      this.valuesMatch(row[pkColumnIndex], pkValue)
    );
    if (recordIndex === -1) return null;

    const record = data[recordIndex];
    let value = record[columnIndex];
    // value === "" の場合は value = 0 として扱う
    if (value === "" || value === null) {
      value = 0;
    } else if (typeof value !== "number") {
      throw new Error(`Column "${columnName}" is not a number`);
    }

    const newValue = value + increment;

    this.getSheet().getRange(recordIndex + this.ROW_INDEX_CONVERT_NUM, columnIndex + 1).setValue(newValue);

    return newValue;
  }

  /**
   * 指定したシートの指定したプライマリキーに基づいて行を更新します。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {Object} rowValues - 更新する行のデータ。キーは列名、値はセルの値です。
   * @returns {Array<Object>} 更新された行のデータ。一致する行がない場合は空の配列。
   */
  updateByPk(pkColumnName: string, pkValue: any, rowValues: object): Array<object> {
    const pkColumnIndex = this.getIndexByColumnName(pkColumnName);
    const data = this.getDataset();

    const recordIndex = data.findIndex((row) =>
      this.valuesMatch(row[pkColumnIndex], pkValue)
    );
    if (recordIndex === -1) return [];

    const record = data[recordIndex];

    const columnIndexMap = this.getColumnIndexMap();
    const newRecord = this.createRecord(rowValues, columnIndexMap);

    this.getSheet()
      .getRange(recordIndex + this.ROW_INDEX_CONVERT_NUM, 1, 1, record.length)
      .setValues([newRecord]);

    return this.convertRecordToObj([newRecord]);
  }

  /**
   * 指定したシートの指定したプライマリキーに基づいて、特定の列の値を更新します。
   * @param {string} pkColumnName - プライマリキーとして使用する列名。
   * @param {*} pkValue - 検索するプライマリキーの値。
   * @param {string} columnName - 更新する列の名前。
   * @param {*} value - 更新する列の値。
   * @returns {Array<Object>} 更新された行のデータ。一致する行がない場合は空の配列。
   */
  updateItemByPk(pkColumnName: string, pkValue: any, columnName: string, value: any): Array<object> {
    const pkColumnIndex = this.getIndexByColumnName(pkColumnName);
    const data = this.getDataset();

    const recordIndex = data.findIndex((row) =>
      this.valuesMatch(row[pkColumnIndex], pkValue)
    );
    if (recordIndex === -1) return [];

    const record = data[recordIndex];
    const columnIndex = this.getIndexByColumnName(columnName);
    record[columnIndex] = value;

    this.getSheet()
      .getRange(recordIndex + this.ROW_INDEX_CONVERT_NUM, 1, 1, record.length)
      .setValues([record]);

    return this.convertRecordToObj([record]);
  }

  /**
   * 指定したシートの指定した列の値に基づいて行を更新します。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {String} columnName - 更新する列の名前。
   * @param {*} value - 更新する列の値。
   * @returns {Array<Object>} 更新された行のデータ。一致する行がない場合は空の配列。
   */
  updateItemByColumns(criteria: object, columnName: string, value: any): Array<object> {
    const columnIndexMap = this.getColumnIndexMap();
    const data = this.getDataset();

    const newRecords = data
      .map((record, index) => {
        const flag = Object.entries(criteria).every(([columnName, value]) => {
          const index = this.getIndexByColumnName(columnName);
          return this.valuesMatch(record[index], value);
        });

        if (flag) {
          const columnIndex = columnIndexMap[columnName];
          record[columnIndex] = value;

          this.getSheet()
            .getRange(index + this.ROW_INDEX_CONVERT_NUM, 1, 1, record.length)
            .setValues([record]);
          return record;
        }
      })
      .filter(Boolean);

    if (newRecords) return [];

    return this.convertRecordToObj(newRecords);
  }

  /**
   * 指定したシートの指定した列の値に基づいて行を更新します。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {Object} columnValues - 更新対象となるカラムと値を持つオブジェクト。
   * @returns {Array<Object>} 更新された行のデータ。一致する行がない場合は空の配列。
   */
  updateItemsByColumns(criteria: object, columnValues: object): Array<object> {
    const columnIndexMap = this.getColumnIndexMap();
    const data = this.getDataset();

    const newRecords = data
      .map((record, index) => {
        const flag = Object.entries(criteria).every(([columnName, value]) => {
          const index = this.getIndexByColumnName(columnName);
          return this.valuesMatch(record[index], value);
        });

        if (flag) {
          Object.entries(columnValues).forEach(([columnName, value]) => {
            const columnIndex = columnIndexMap[columnName];
            record[columnIndex] = value;
          });

          this.getSheet()
            .getRange(index + this.ROW_INDEX_CONVERT_NUM, 1, 1, record.length)
            .setValues([record]);
          return record;
        }
      })
      .filter(Boolean);

    if (newRecords) return [];

    return this.convertRecordToObj(newRecords);
  }

  // the following methods are private
  /**
   * レコードの配列をオブジェクトの配列に変換します。
   * @param {Array<Array<*>>} records - レコードの配列。
   * @returns {Array<Object>} レコードの配列をオブジェクトの配列に変換したもの。
   */
  private convertRecordToObj(records: Array<Array<any>>): Array<object> {
    const columnIndexMap = this.getColumnIndexMap();
    return records.map((record) => {
      const obj = {};
      Object.entries(columnIndexMap).forEach(([column, index]) => {
        obj[column] = record[index];
      });
      return obj;
    });
  }

  /**
   * カラム名に基づいてインデックスを取得します。
   * @private
   * @param {string} columnName - カラム名
   * @returns {number} カラムのインデックス
   * @throws {Error} 無効なカラム名の場合
   */
  private getIndexByColumnName(columnName: string): number {
    const columnIndexMap = this.getColumnIndexMap();
    const index = columnIndexMap[columnName];
    if (index === undefined)
      throw new Error(`Invalid column name: ${columnName}`);
    return parseInt(index);
  }

  /**
   * シートを取得します。
   * @private
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
   */
  private getSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    return this.sheet;
  }

  /**
   * データセットを取得します。
   * @private
   * @returns {Array<Array>} シートからデータを取得し、配列形式で返す
   */
  private getDataset(): Array<Array<any>> {
    return this.getSheet().getDataRange().getValues().slice(1);
  }

  /**
   * カラムインデックスマップを取得します。
   * @private
   * @returns {Object} カラム名とインデックスのマップ
   */
  private getColumnIndexMap(): object {
    return this.columnIndexMap;
  }

  /**
   * ヘッダー行からカラムインデックスマップを作成します。
   * @private
   * @param {Array} headerRow - ヘッダー行のデータ
   * @returns {{[key: string]: number}} カラム名とインデックスのマップ
   */
  private createColumnIndexMap(headerRow: Array<any>): {[key: string]: number} {
    const columnIndexMap = {};
    headerRow.forEach((column, index) => {
      columnIndexMap[column] = index;
    });
    return columnIndexMap;
  }

  /**
   * 行の値とカラムインデックスマップを使用して、新しいレコードを作成します。
   * @private
   * @param {Object} rowValues - 行の値
   * @param {Object} columnIndexMap - カラム名とインデックスのマップ
   * @returns {Array} 新しいレコード
   */
  private createRecord(rowValues: object, columnIndexMap: object): Array<any> {
    const newRecord = new Array(Object.keys(columnIndexMap).length).fill("");
    Object.entries(rowValues).forEach(([column, value]) => {
      const colIndex = columnIndexMap[column];
      if (colIndex !== undefined) {
        newRecord[colIndex] = value;
      }
    });

    return newRecord;
  }

  /**
   * データをソートします。
   * @private
   * @param {Array} records - ソートするデータ
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。
   * @returns {Array} ソートされたデータ
   * @throws {Error} 無効なソートカラム名の場合
   */
  private sortData(records: Array<any>, sortBy: {column: string, order: "ASC" | "DESC"}): Array<any> {
    if (!sortBy) return records;
    const { column, order } = sortBy;
    const index = this.getIndexByColumnName(column)

    if (index === undefined) {
      throw new Error(`Invalid column name forsorting: ${column}`);
    }

    return records.sort((rowA, rowB) => {
      const valueA = rowA[index];
      const valueB = rowB[index];
      const comparison = valueA < valueB ? -1 : valueA > valueB ? 1 : 0;

      return order === "ASC" ? comparison : -comparison;
    });
  }

  /**
   * 2つの値が一致するかどうかを判断します。
   * @private
   * @param {*} value1 - 値1
   * @param {*} value2 - 値2
   * @returns {boolean} 一致する場合は true、それ以外の場合は false
   */
  private valuesMatch(value1: any, value2: any): boolean {
    // if both values are null, they match
    return String(value1) === String(value2);
  }
}