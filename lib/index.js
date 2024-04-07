"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
var XLSX = require("xlsx");
var XLSXStyle = require("xlsx-js-style");
var FileSaver = require("file-saver");
function convertDataToExcel(_a) {
    var _b;
    var data = _a.data, _c = _a.fileExtension, fileExtension = _c === void 0 ? ".xlsx" : _c, _d = _a.startCellOfData, startCellOfData = _d === void 0 ? "A1" : _d, colsGroup = _a.colsGroup, _e = _a.sheetName, sheetName = _e === void 0 ? "Sheet 1" : _e, colWidths = _a.colWidths, rowHeights = _a.rowHeights, _f = _a.fileName, fileName = _f === void 0 ? "example" : _f, isBordered = _a.isBordered, styles = _a.styles;
    var fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
    var options = {
        origin: startCellOfData,
    };
    var ws = XLSX.utils.json_to_sheet(data, options);
    if (colsGroup === null || colsGroup === void 0 ? void 0 : colsGroup.length) {
        colsGroup === null || colsGroup === void 0 ? void 0 : colsGroup.forEach(function (colGroup) {
            ws["!merges"] = __spreadArray(__spreadArray([], (ws["!merges"] || []), true), [
                {
                    s: { c: colGroup.colStart, r: colGroup.rowStart },
                    e: { c: colGroup.colEnd, r: colGroup.rowEnd },
                },
            ], false);
            XLSX.utils.sheet_add_aoa(ws, [[colGroup.content]], {
                origin: colGroup.origin,
            });
        });
    }
    if (styles === null || styles === void 0 ? void 0 : styles.length) {
        var _loop_1 = function () {
            var cell = XLSX.utils.decode_cell(i);
            styles.forEach(function (style) {
                var checkCellAvaiable = true;
                if (!style.rowEnd) {
                    if (cell.r < style.rowStart) {
                        checkCellAvaiable = false;
                    }
                }
                else {
                    if (cell.r < style.rowStart || cell.r > style.rowEnd) {
                        checkCellAvaiable = false;
                    }
                }
                if (!style.colEnd) {
                    if (cell.c < style.colStart) {
                        checkCellAvaiable = false;
                    }
                }
                else {
                    if (cell.c < style.colStart || cell.c > style.colEnd) {
                        checkCellAvaiable = false;
                    }
                }
                if (checkCellAvaiable) {
                    ws[i].s = style.style;
                    ws[i].t = style.type || "s";
                    ws[i].z = style.format;
                }
            });
        };
        for (var i in ws) {
            _loop_1();
        }
    }
    if (isBordered) {
        for (var i in ws) {
            var cell = XLSX.utils.decode_cell(i);
            var rowOfHeader = Number(startCellOfData.split("").pop());
            if (cell.r >= rowOfHeader - 1) {
                ws[i].s = __assign(__assign({}, ws[i].s), { border: {
                        right: { style: "thin" },
                        left: { style: "thin" },
                        top: { style: "thin" },
                        bottom: { style: "thin" },
                    } });
            }
        }
    }
    if (colWidths === null || colWidths === void 0 ? void 0 : colWidths.length) {
        var wscols = colWidths.map(function (colWidth) { return ({
            wch: colWidth,
        }); });
        ws["!cols"] = wscols;
    }
    if (rowHeights === null || rowHeights === void 0 ? void 0 : rowHeights.length) {
        var wsrows = rowHeights.map(function (rowHeight) { return ({
            hpt: rowHeight,
        }); });
        ws["!rows"] = wsrows;
    }
    var wb = { Sheets: (_b = {}, _b[sheetName] = ws, _b), SheetNames: [sheetName] };
    var excelBuffer = XLSXStyle.write(wb, {
        bookType: "xlsx",
        type: "array",
    });
    var dataConverted = new Blob([excelBuffer], { type: fileType });
    FileSaver.saveAs(dataConverted, fileName + fileExtension);
    return dataConverted;
}
exports.default = convertDataToExcel;
