const Excel = require('exceljs');
exports.createExcel = async (options) => {
    let _columns, _path, _rows, _sheet_name, sheet, workbook;
    _sheet_name = options.sheet_name;
    _rows = options.rows;
    _columns = options.columns;
    _path = options.path || 'tmp/default.xlsx';
    workbook = new Excel.Workbook();
    sheet = workbook.addWorksheet(_sheet_name);
    sheet.columns = _columns; sheet.addRows(_rows);
    sheet.getRow(1).font = { bold: true };
    return workbook.xlsx.writeFile(_path).then(err=>{
        if (err) {
            return err;
        } else {
            return _path;
        }
    });
};