var _ = require('lodash');
var XLSX = require('xlsx');
var worksheet = (function () {
    function worksheet(name) {
        this.headerColumns = [];
        this.rows = [];
        this.C = 0;
        this.R = 0;
        this.range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
        this.data = {};
        this.name = name;
    }
    worksheet.prototype.datenum = function (v, date1904) {
        if (date1904)
            v += 1462;
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    };
    worksheet.prototype.setHeader = function (arr) {
        var _this = this;
        _.each(arr, function (item) {
            _this.headerColumns.push(item);
        });
        _.each(this.headerColumns, function (col) {
            _this.setCell(_this.R, _this.C, { v: col });
            _this.C++;
        });
        this.R++;
        this.C = 0;
    };
    worksheet.prototype.addRow = function (row) {
        var _this = this;
        if (this.headerColumns) {
            _.each(this.headerColumns, function (col) {
                var item = row[col];
                var cell;
                if (typeof item === 'object' && item.v !== "undefined") {
                    cell = item;
                }
                else {
                    cell = { v: item };
                }
                _this.setCell(_this.R, _this.C, cell);
                _this.C++;
            });
            this.R++;
            this.C = 0;
        }
    };
    worksheet.prototype.setCell = function (R, C, cell) {
        var ws = {};
        if (this.range.s.r > R)
            this.range.s.r = R;
        if (this.range.s.c > C)
            this.range.s.c = C;
        if (this.range.e.r < R)
            this.range.e.r = R;
        if (this.range.e.c < C)
            this.range.e.c = C;
        if (cell.v != null) {
            var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
            if (typeof cell.v === 'number')
                cell.t = 'n';
            else if (typeof cell.v === 'boolean')
                cell.t = 'b';
            else if (cell.v instanceof Date) {
                cell.t = 'n';
                if (cell.z == null) {
                    cell.z = XLSX.SSF._table[14];
                }
                cell.v = this.datenum(cell.v);
            }
            else {
                cell.t = 's';
            }
            this.data[cell_ref] = cell;
        }
    };
    worksheet.prototype.write = function () {
        if (this.range.s.c < 10000000) {
            this.data['!ref'] = XLSX.utils.encode_range(this.range);
        }
    };
    return worksheet;
})();
var workbook = (function () {
    function workbook() {
        this.worksheets = [];
    }
    workbook.prototype.addSheet = function (worksheet) {
        this.worksheets.push(worksheet);
    };
    workbook.prototype.write = function () {
        var wb = {
            SheetNames: [],
            Sheets: {}
        };
        _.each(this.worksheets, function (sheet) {
            sheet.write();
            wb.SheetNames.push(sheet.name);
            wb.Sheets[sheet.name] = sheet.data;
        });
        var wopts = { bookType: 'xlsx', bookSST: false, type: 'binary' };
        var output = XLSX.write(wb, wopts);
        return output;
    };
    workbook.prototype.download = function (res, filename) {
        var wbout = this.write();
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.attachment(filename);
        res.send(new Buffer(wbout, 'binary'));
    };
    return workbook;
})();
var xslx_tools = (function () {
    function xslx_tools() {
    }
    xslx_tools.prototype.xlsx_tools = function () {
    };
    xslx_tools.workbook = function () {
        return new workbook;
    };
    xslx_tools.worksheet = function (name) {
        return new worksheet(name);
    };
    xslx_tools.init = function (m) {
        return new xslx_tools();
    };
    return xslx_tools;
})();
module.exports = xslx_tools;
