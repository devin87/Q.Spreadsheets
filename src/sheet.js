/// <reference path="core.js" />
/*
* sheet.js
* update: 2015/11/19 17:21
*/
(function (window, undefined) {
    "use strict";

    var extend = Q.extend,
        clone = Q.clone,
        toMap = Q.toMap,

        factory = Q.factory,

        SS = Q.SS;

    /*
    * Sheet 对象
    */
    function Sheet(data) {
        this.init(data);
    }

    factory(Sheet).extend({
        //初始化
        init: function (data) {
            return this.parse(data);
        },
        //重置
        reset: function () {
            return this.parse({});
        },
        //解析数据
        parse: function (data) {
            data = data || {};

            var self = this;

            self.name = data.name;
            self.tabColor = data.tabColor;
            self.colWidth = data.colWidth || 80;
            self.rowHeight = data.rowHeight || 20;

            self.displayFormula = !!data.displayFormula;
            self.displayGridlines = !!data.displayGridlines;
            self.displayHeadings = !!data.displayHeadings;

            self.firstRow = data.firstRow || 0;
            self.firstCol = data.firstCol || 0;
            self.lastRow = data.lastRow || 0;
            self.lastCol = data.lastCol || 0;
            self.scrollRow = data.scrollRow || 0;
            self.scrollCol = data.scrollCol || 0;

            self.freezeRow = data.freezeRow;
            self.freezeCol = data.freezeCol;

            self.isFreeze = !!data.isFreeze;
            self.isHidden = !!data.isHidden;

            self.style = data.style;
            self.rowStyles = data.rowStyles || {};
            self.colStyles = data.colStyles || {};
            self.hiddenRows = toMap(data.hiddenRows || [], true);
            self.hiddenCols = toMap(data.hiddenCols || [], true);
            self.merges = data.merges || [];
            self.links = data.links || [];
            self.comments = data.comments || [];
            self.richTexts = data.richTexts || [];

            //----------------- parse cell data -----------------
            var cells = [],

                dataCells = data.cells || [],
                count = dataCells.length,
                i = 0,

                dataCell,
                rowIndex;

            for (; i < count; i++) {
                //dataCell => [row,col,type,value,format,style]
                dataCell = dataCells[i];
                rowIndex = dataCell[0];

                if (!cells[rowIndex]) cells[rowIndex] = [];
                cells[rowIndex][dataCell[1]] = { type: dataCell[2], value: dataCell[3], format: dataCell[4], style: dataCell[5] };
            }

            self.cells = cells;

            return self;
        },
        //克隆工作表
        clone: function () {
            var self = this,
                sheet = new Sheet;

            Sheet.PROPS.forEach(function (prop) {
                sheet[prop] = clone(self[prop]);
            });

            return sheet;
        },

        getData: function () {
            var self = this, data = {}, dataCells = [];

            Sheet.PROPS.forEach(function (prop) {
                data[prop] = self[prop];
            });

            data.cells.forEach(function (row, i) {
                if (!row) return;

                row.forEach(function (cell, j) {
                    if (cell) dataCells.push([i, j, cell.type, cell.value, cell.format, cell.style]);
                });
            });

            data.cells = dataCells;
        }
    });

    var DEF_SHEET = new Sheet;

    extend(Sheet, {
        DEF: DEF_SHEET,
        PROPS: Object.keys(DEF_SHEET)
    });

    SS.Sheet = Sheet;

})(window);