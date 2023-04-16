/**
 * SpreadsheetDBを作成します。
 * @returns {SpreadsheetDB} アクティブなスプレッドシートのSpreadsheetDB。
 * @throws {Error} スプレッドシートが見つからない場合。
 */
function create(): SpreadsheetDB {
    try {
        return new SpreadsheetDB("");
    } catch (e) {
        Logger.log(e);
        throw new Error("スプレッドシートが見つかりません");
    }
}

/**
 * SpreadsheetDBを作成します。
 * @param {string} spreadsheetId スプレッドシートのID。
 * @returns {SpreadsheetDB} 指定したスプレッドシートのSpreadsheetDB。
 */
function createById(spreadsheetId: string): SpreadsheetDB {

    try {
        return new SpreadsheetDB(spreadsheetId);
    } catch (e) {
        Logger.log(e);
        throw new Error("スプレッドシートが見つかりません");
    }
}