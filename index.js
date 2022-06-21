"use strict";
exports.__esModule = true;
var xlsx_1 = require("xlsx");
var JsonToXlsx = /** @class */ (function () {
    function JsonToXlsx(data, sheetName) {
        if (sheetName === void 0) { sheetName = 'Sheet1'; }
        this.originData = data;
        this.ws = xlsx_1.utils.json_to_sheet(this.originData);
        this.wb = xlsx_1.utils.book_new();
        xlsx_1.utils.book_append_sheet(this.wb, this.ws, sheetName);
    }
    JsonToXlsx.prototype.replaceHeader = function (obj) {
        var range = xlsx_1.utils.decode_range(this.ws['!ref']);
        for (var i = range.s.c; i <= range.e.c; i++) {
            var h = xlsx_1.utils.encode_col(i) + '1';
            obj[this.ws[h].v] && (this.ws[h].v = obj[this.ws[h].v]);
        }
        return this;
    };
    JsonToXlsx.prototype.download = function (fileName) {
        fileName !== null && fileName !== void 0 ? fileName : (fileName = 'excel' + Date.now());
        return (0, xlsx_1.writeFile)(this.wb, fileName + ".xlsx");
    };
    return JsonToXlsx;
}());
exports["default"] = JsonToXlsx;
