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
    const sheetTable = this.getSheetTable(sheetName);
    return sheetTable.insert(rowValues);
  }

  /**
   * 指定したシートに複数の行を追加します。
   * @param {string} sheetName - 行を追加するシートの名前。
   * @param {Array<Object>} rowValuesArray - 追加する行のデータの配列。各要素は、キーが列名、値がセルの値のオブジェクトです。
   * @returns {number} 追加された行数。
   */
  insertAll(sheetName, rowValuesArray) {
    const sheetTable = this.getSheetTable(sheetName);
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
    const sheetTable = this.getSheetTable(sheetName);
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
    const sheetTable = this.getSheetTable(sheetName);
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
    const sheetTable = this.getSheetTable(sheetName);
    return sheetTable.selectByColumnSorted(columnName, value, sortBy);
  }

  /**
   * 指定したシートから、指定した複数の列の値に基づいて行を検索します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
   */
  selectByColumns(sheetName, criteria) {
    const sheetTable = this.getSheetTable(sheetName);
    return sheetTable.selectByColumns(criteria);
  }

  /**
   * シートから、全てのデータを取得します。
   * @param {string} sheetName - 行を検索するシートの名前。
   * @returns {Array<Object>|null} シートの全行のデータの配列。
   */
  selectAll(sheetName) {
    const sheetTable = this.getSheetTable(sheetName);
    return sheetTable.selectAll();
  }

  selectMax(sheetName, columnName) {
    const sheetTable = this.getSheetTable(sheetName);
    return sheetTable.selectMax(columnName);
  }

  selectByPkAndIncrement(sheetName, pkColumnName, pkValue, columnName, increment) {
    const sheetTable = this.getSheetTable(sheetName);
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
    const sheetTable = this.getSheetTable(sheetName);
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
    const sheetTable = this.getSheetTable(sheetName);
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
    const sheetTable = this.getSheetTable(sheetName);
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
    const sheetTable = this.getSheetTable(sheetName);
    return sheetTable.updateItemsByColumns(criteria, columnValues);
  }

  getSheetTable(sheetName) {
    if (this.sheetMap[sheetName]) {
      return this.sheetMap[sheetName];
    }
    const sheet = new SheetTable(this.spreadsheet, sheetName);
    this.sheetMap[sheetName] = sheet;
    return sheet;
  }
}

/**
 * SheetTable クラスは、Google スプレッドシートのシートを操作するためのクラスです。
 */
class SheetTable {
  /**
   * SheetTable クラスのインスタンスを作成します。@param {Object} spreadsheet - Google スプレッドシートのインスタンス。
   * @param {string} sheetName - シートの名前。
   */
  constructor(spreadsheet, sheetName) {
    this.sheetName = sheetName;
    try {
      this.sheet = spreadsheet.getSheetByName(this.sheetName);
    } catch (e) {
      console.log(e);
      throw new Error("シートが見つかりません");
    }
    this.headers = this.sheet
      .getRange(1, 1, 1, this.sheet.getLastColumn())
      .getValues()[0];
    this.columnIndexMap = this.createColumnIndexMap(this.headers);
    // row index => row number
    this.ROW_INDEX_CONVERT_NUM = 2;
  }

  /**
   * シートに行を追加します。
   * @param {Object} rowValues - 追加する行のデータ。キーは列名、値はセルの値です。
   * @returns {number} 追加された行数。
   */
  insert(rowValues) {
    const record = this.createRecord(rowValues, this.getColumnIndexMap());
    this.getSheet().appendRow(record);
    return 1;
  }

  /**
   * シートに複数の行を追加します。
   * @param {Array<Object>} rowValuesArray - 追加する行のデータの配列。各要素は、キーが列名、値がセルの値のオブジェクトです。
   * @returns {number} 追加された行数。
   */
  insertAll(rowValuesArray) {
    const records = rowValuesArray.map((rowValues) =>
      this.createRecord(rowValues, this.getColumnIndexMap())
    );
    this.getSheet()
      .getRange(
        this.getLastRow() + this.ROW_INDEX_CONVERT_NUM,
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
   * @returns {Array<Object>|} 検索に一致した行のデータ。一致する行がない場合は null。
   */
  selectByPk(pkColumnName, pkValue) {
    const index = this.getIndexByColumnName(pkColumnName);
    const data = this.getDataset();
    const record = data.find((row) => this.valuesMatch(row[index], pkValue));
    if (!record) return null;

    return this.convertRecordToObj([record]);
  }

  /**
   * シートから、指定した列の値に基づいて行を検索します。
   * @param {string} columnName - 検索に使用する列名。
   * @param {*} value - 検索する列の値。
   * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
   */
  selectByColumn(columnName, value) {
    const index = this.getIndexByColumnName(columnName);
    const data = this.getDataset();
    const records = data.filter((row) => {
      return this.valuesMatch(row[index], value);
    });

    if (records.length === 0) return null;

    return this.convertRecordToObj(records);
  }

  /**
   * シートから、指定した列の値に基づいて行を検索し、結果を指定した条件でソートします。
   * @param {string} columnName - 検索に使用する列名。
   * @param {*} value - 検索する列の値。
   * @param {{column: string, order: "ASC" | "DESC"}} sortBy - 結果をソートするための条件。キーが列名、order は "ASC"（昇順）または "DESC"（降順）です。
   * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
   */
  selectByColumnSorted(columnName, value, sortBy) {
    const records = this.selectByColumn(columnName, value);
    if (records.length === 0) return null;

    const sortedRecords = this.sortData(
      records,
      sortBy,
      this.getColumnIndexMap()
    );
    return this.convertRecordToObj(sortedRecords);
  }

  /**
   * シートから、指定した複数の列の値に基づいて行を検索します。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
   */
  selectByColumns(criteria) {
    const data = this.getDataset();
    const records = data.filter(
      (row) =>
        Object.keys(criteria).length === 0 ||
        Object.entries(criteria).every(([columnName, value]) => {
          const index = this.getIndexByColumnName(columnName);
          return this.valuesMatch(row[index], value);
        })
    );

    if (records.length === 0) return null;

    return this.convertRecordToObj(records);
  }

  /**
   * シートから、全てのデータを取得します。
   * @return {Array<Object>} シートの全行のデータの配列。
   */
  selectAll() {
    const data = this.getDataset();
    return this.convertRecordToObj(data);
  }

  /**
   * シートから、指定した列の最大値を取得します。
   * @param {string} columnName - 最大値を取得する列名。
   * @returns {number|null} 指定した列の最大値。列が存在しない場合は null。
   */
  selectMax(columnName) {
    if (typeof columnName !== "string" || columnName === "") {
      throw new Error("Invalid columnName");
    }

    const index = this.getIndexByColumnName(columnName);
    if (index === undefined) {
      throw new Error(`Column "${columnName}" not found`);
    }

    const data = this.getDataset();
    if (data.length === 0) {
      return null;
    }

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
   * @returns {Object|null} インクリメント後の行のデータ。行が存在しない場合は null。
   */
  selectByPkAndIncrement(pkColumnName, pkValue, columnName, increment = 1) {
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
    const value = record[columnIndex];
    if (typeof value !== "number") {
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
   * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
   */
  updateByPk(pkColumnName, pkValue, rowValues) {
    const pkColumnIndex = this.getIndexByColumnName(pkColumnName);
    const data = this.getDataset();

    const recordIndex = data.findIndex((row) =>
      this.valuesMatch(row[pkColumnIndex], pkValue)
    );
    if (recordIndex === -1) return null;

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
   * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
   */
  updateItemByPk(pkColumnName, pkValue, columnName, value) {
    const pkColumnIndex = this.getIndexByColumnName(pkColumnName);
    const data = this.getDataset();

    const recordIndex = data.findIndex((row) =>
      this.valuesMatch(row[pkColumnIndex], pkValue)
    );
    if (recordIndex === -1) return null;

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
   * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
   */
  updateItemByColumns(criteria, columnName, value) {
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

    if (newRecords.length === 0) return null;

    return this.convertRecordToObj(newRecords);
  }

  /**
   * 指定したシートの指定した列の値に基づいて行を更新します。
   * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
   * @param {Object} columnValues - 更新対象となるカラムと値を持つオブジェクト。
   * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
   */
  updateItemsByColumns(criteria, columnValues) {
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

    if (newRecords.length === 0) return null;

    return this.convertRecordToObj(newRecords);
  }

  /**
   * 対象シート内の複数のセルを一度のリクエストで更新します。
   * @param {Object} cellUpdates - セルのアドレス（例："A1"）をキーとし、セルの新しい値を値とするオブジェクト。
   * @returns {Boolean} - 操作が成功した場合は true を、そうでない場合は false を返します。
   * @description このメソッドは、Google sheet api のバッチ更新機能を使用しています。
   */
  updateCells(cellUpdates) {
    if (!cellUpdates || typeof cellUpdates !== "object") {
      throw new Error(
        "Invalid input for cellUpdates, must be an object with cell addresses as keys and new values as values."
      );
    }

    const sheet = this.getSheet();
    const requests = [];

    for (const [cellAddress, newValue] of Object.entries(cellUpdates)) {
      requests.push({
        updateCells: {
          range: {
            sheetId: sheet.getSheetId(),
            startRowIndex: sheet.getRange(cellAddress).getRowIndex() - 1,
            endRowIndex: sheet.getRange(cellAddress).getRowIndex(),
            startColumnIndex: sheet.getRange(cellAddress).getColumn() - 1,
            endColumnIndex: sheet.getRange(cellAddress).getColumn(),
          },
          rows: [
            {
              values: [
                {
                  userEnteredValue: { stringValue: newValue },
                },
              ],
            },
          ],
          fields: "userEnteredValue",
        },
      });
    }

    try {
      const response = Sheets.Spreadsheets.batchUpdate(
        { requests: requests },
        sheet.getParent().getId()
      );
      return response && response.replies.length === requests.length;
    } catch (e) {
      console.log(e);
      return false;
    }
  }

  // the following methods are private

  /**
   * レコードの配列をオブジェクトの配列に変換します。
   * @param {Array<Array<*>>} records - レコードの配列。
   * @returns {Array<Object>} レコードの配列をオブジェクトの配列に変換したもの。
   */
  convertRecordToObj(records) {
    const columnIndexMap = this.getColumnIndexMap();
    return records.map((record) => {
      const obj = {};
      Object.entries(columnIndexMap).forEach(([column, index]) => {
        obj[column] = record[index];
      });
      return obj;
    });
  }

  getIndexByColumnName(columnName) {
    const columnIndexMap = this.getColumnIndexMap();
    const index = columnIndexMap[columnName];
    if (index === undefined)
      throw new Error(`Invalid column name: ${columnName}`);
    return parseInt(index);
  }

  getSheet() {
    return this.sheet;
  }

  getDataset() {
    return this.getSheet().getDataRange().getValues().slice(1);
  }

  getColumnIndexMap() {
    return this.columnIndexMap;
  }

  createColumnIndexMap(headerRow) {
    const columnIndexMap = {};
    headerRow.forEach((column, index) => {
      columnIndexMap[column] = index;
    });
    return columnIndexMap;
  }

  createRecord(rowValues, columnIndexMap) {
    const newRecord = new Array(Object.keys(columnIndexMap).length).fill("");
    Object.entries(rowValues).forEach(([column, value]) => {
      const colIndex = columnIndexMap[column];
      if (colIndex !== undefined) {
        newRecord[colIndex] = value;
      }
    });

    return newRecord;
  }

  sortData(records, sortBy, columnIndexMap) {
    if (!sortBy) return records;
    const { column, order } = sortBy;
    const index = columnIndexMap[column];

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

  valuesMatch(value1, value2) {
    // if both values are null, they match
    return String(value1) === String(value2);
  }
}