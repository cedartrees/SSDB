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
      this.columnIndexMap = this.createColumnIndexMap_(this.headers);
      // row index => row number
      this.ROW_INDEX_CONVERT_NUM = 2;
    }
  
    /**
     * シートに行を追加します。
     * @param {Object} rowValues - 追加する行のデータ。キーは列名、値はセルの値です。
     * @returns {number} 追加された行数。
     */
    insert(rowValues) {
      const record = this.createRecord_(rowValues, this.getColumnIndexMap_());
      this.getSheet_().appendRow(record);
      return 1;
    }
  
    /**
     * シートに複数の行を追加します。
     * @param {Array<Object>} rowValuesArray - 追加する行のデータの配列。各要素は、キーが列名、値がセルの値のオブジェクトです。
     * @returns {number} 追加された行数。
     */
    insertAll(rowValuesArray) {
      const records = rowValuesArray.map((rowValues) =>
        this.createRecord_(rowValues, this.getColumnIndexMap_())
      );
      this.getSheet_()
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
      const index = this.getIndexByColumnName_(pkColumnName);
      const data = this.getDataset_();
      const record = data.find((row) => this.valuesMatch_(row[index], pkValue));
      if (!record) return null;
  
      return this.convertRecordToObj_([record]);
    }
  
    /**
     * シートから、指定した列の値に基づいて行を検索します。
     * @param {string} columnName - 検索に使用する列名。
     * @param {*} value - 検索する列の値。
     * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
     */
    selectByColumn(columnName, value) {
      const index = this.getIndexByColumnName_(columnName);
      const data = this.getDataset_();
      const records = data.filter((row) => {
        return this.valuesMatch_(row[index], value);
      });
  
      if (records.length === 0) return null;
  
      return this.convertRecordToObj_(records);
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
  
      const sortedRecords = this.sortData_(
        records,
        sortBy,
        this.getColumnIndexMap_()
      );
      return this.convertRecordToObj_(sortedRecords);
    }
  
    /**
     * シートから、指定した複数の列の値に基づいて行を検索します。
     * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
     * @returns {Array<Object>|null} 検索に一致した行のデータの配列。一致する行がない場合は null。
     */
    selectByColumns(criteria) {
      const data = this.getDataset_();
      const records = data.filter(
        (row) =>
          Object.keys(criteria).length === 0 ||
          Object.entries(criteria).every(([columnName, value]) => {
            const index = this.getIndexByColumnName_(columnName);
            return this.valuesMatch_(row[index], value);
          })
      );
  
      if (records.length === 0) return null;
  
      return this.convertRecordToObj_(records);
    }
  
    /**
     * シートから、全てのデータを取得します。
     * @return {Array<Object>} シートの全行のデータの配列。
     */
    selectAll() {
      const data = this.getDataset_();
      return this.convertRecordToObj_(data);
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
  
      const index = this.getIndexByColumnName_(columnName);
      if (index === undefined) {
        throw new Error(`Column "${columnName}" not found`);
      }
  
      const data = this.getDataset_();
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
    
      const pkColumnIndex = this.getIndexByColumnName_(pkColumnName);
      if (pkColumnIndex === undefined) {
        throw new Error(`Column "${pkColumnName}" not found`);
      }
    
      const columnIndex = this.getIndexByColumnName_(columnName);
      if (columnIndex === undefined) {
        throw new Error(`Column "${columnName}" not found`);
      }
    
      const data = this.getDataset_();
    
      const recordIndex = data.findIndex((row) =>
        this.valuesMatch_(row[pkColumnIndex], pkValue)
      );
      if (recordIndex === -1) return null;
    
      const record = data[recordIndex];
      const value = record[columnIndex];
      if (typeof value !== "number") {
        throw new Error(`Column "${columnName}" is not a number`);
      }
    
      const newValue = value + increment;
    
      this.getSheet_().getRange(recordIndex + this.ROW_INDEX_CONVERT_NUM, columnIndex + 1).setValue(newValue);
    
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
      const pkColumnIndex = this.getIndexByColumnName_(pkColumnName);
      const data = this.getDataset_();
  
      const recordIndex = data.findIndex((row) =>
        this.valuesMatch_(row[pkColumnIndex], pkValue)
      );
      if (recordIndex === -1) return null;
  
      const record = data[recordIndex];
  
      const columnIndexMap = this.getColumnIndexMap_();
      const newRecord = this.createRecord_(rowValues, columnIndexMap);
  
      this.getSheet_()
        .getRange(recordIndex + this.ROW_INDEX_CONVERT_NUM, 1, 1, record.length)
        .setValues([newRecord]);
  
      return this.convertRecordToObj_([newRecord]);
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
      const pkColumnIndex = this.getIndexByColumnName_(pkColumnName);
      const data = this.getDataset_();
  
      const recordIndex = data.findIndex((row) =>
        this.valuesMatch_(row[pkColumnIndex], pkValue)
      );
      if (recordIndex === -1) return null;
  
      const record = data[recordIndex];
      const columnIndex = this.getIndexByColumnName_(columnName);
      record[columnIndex] = value;
  
      this.getSheet_()
        .getRange(recordIndex + this.ROW_INDEX_CONVERT_NUM, 1, 1, record.length)
        .setValues([record]);
  
      return this.convertRecordToObj_([record]);
    }
  
    /**
     * 指定したシートの指定した列の値に基づいて行を更新します。
     * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
     * @param {String} columnName - 更新する列の名前。
     * @param {*} value - 更新する列の値。
     * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
     */
    updateItemByColumns(criteria, columnName, value) {
      const columnIndexMap = this.getColumnIndexMap_();
      const data = this.getDataset_();
  
      const newRecords = data
        .map((record, index) => {
          const flag = Object.entries(criteria).every(([columnName, value]) => {
            const index = this.getIndexByColumnName_(columnName);
            return this.valuesMatch_(record[index], value);
          });
  
          if (flag) {
            const columnIndex = columnIndexMap[columnName];
            record[columnIndex] = value;
  
            this.getSheet_()
              .getRange(index + this.ROW_INDEX_CONVERT_NUM, 1, 1, record.length)
              .setValues([record]);
            return record;
          }
        })
        .filter(Boolean);
  
      if (newRecords.length === 0) return null;
  
      return this.convertRecordToObj_(newRecords);
    }
  
    /**
     * 指定したシートの指定した列の値に基づいて行を更新します。
     * @param {Object} criteria - 検索条件となるカラムと値を持つオブジェクト。
     * @param {Object} columnValues - 更新対象となるカラムと値を持つオブジェクト。
     * @returns {Array<Object>|null} 更新された行のデータ。一致する行がない場合は null。
     */
    updateItemsByColumns(criteria, columnValues) {
      const columnIndexMap = this.getColumnIndexMap_();
      const data = this.getDataset_();
  
      const newRecords = data
        .map((record, index) => {
          const flag = Object.entries(criteria).every(([columnName, value]) => {
            const index = this.getIndexByColumnName_(columnName);
            return this.valuesMatch_(record[index], value);
          });
  
          if (flag) {
            Object.entries(columnValues).forEach(([columnName, value]) => {
              const columnIndex = columnIndexMap[columnName];
              record[columnIndex] = value;
            });
  
            this.getSheet_()
              .getRange(index + this.ROW_INDEX_CONVERT_NUM, 1, 1, record.length)
              .setValues([record]);
            return record;
          }
        })
        .filter(Boolean);
  
      if (newRecords.length === 0) return null;
  
      return this.convertRecordToObj_(newRecords);
    }
  
    // the following methods are private
    /**
     * レコードの配列をオブジェクトの配列に変換します。
     * @param {Array<Array<*>>} records - レコードの配列。
     * @returns {Array<Object>} レコードの配列をオブジェクトの配列に変換したもの。
     */
    convertRecordToObj_(records) {
      const columnIndexMap = this.getColumnIndexMap_();
      return records.map((record) => {
        const obj = {};
        Object.entries(columnIndexMap).forEach(([column, index]) => {
          obj[column] = record[index];
        });
        return obj;
      });
    }
  
    getIndexByColumnName_(columnName) {
      const columnIndexMap = this.getColumnIndexMap_();
      const index = columnIndexMap[columnName];
      if (index === undefined)
        throw new Error(`Invalid column name: ${columnName}`);
      return parseInt(index);
    }
  
    getSheet_() {
      return this.sheet;
    }
  
    getDataset_() {
      return this.getSheet_().getDataRange().getValues().slice(1);
    }
  
    getColumnIndexMap_() {
      return this.columnIndexMap;
    }
  
    createColumnIndexMap_(headerRow) {
      const columnIndexMap = {};
      headerRow.forEach((column, index) => {
        columnIndexMap[column] = index;
      });
      return columnIndexMap;
    }
  
    createRecord_(rowValues, columnIndexMap) {
      const newRecord = new Array(Object.keys(columnIndexMap).length).fill("");
      Object.entries(rowValues).forEach(([column, value]) => {
        const colIndex = columnIndexMap[column];
        if (colIndex !== undefined) {
          newRecord[colIndex] = value;
        }
      });
  
      return newRecord;
    }
  
    sortData_(records, sortBy, columnIndexMap) {
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
  
    valuesMatch_(value1, value2) {
      // if both values are null, they match
      return String(value1) === String(value2);
    }
  }