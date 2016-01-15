/// <reference path="core.js" />
/// <reference path="errors.js" />
/// <reference path="lang/zh_CN.js" />
/*
* workbook.js
* author:devin87@qq.com
* update: 2015/12/02 12:17
*/
(function (window, undefined) {
    "use strict";

    var extend = Q.extend,
        clone = Q.clone,
        toMap = Q.toMap,
        toSet = Q.toSet,
        makeArray = Q.makeArray,

        factory = Q.factory,

        SS = Q.SS,
        Errors = SS.Errors,
        Lang = Q.SSLang;

    /*
    * Workbook 对象
    */
    function Workbook(data) {
        var self = this;

        self.init(data);
    }

    factory(Workbook).extend({
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

            self.index = data.index || 0;
            self.title = data.title || "";
            self.subject = data.subject || "";
            self.author = data.author || "";
            self.company = data.company || "";

            self.fonts = data.fonts || [];
            self.styles = data.styles || [];
            self.formats = data.formats || [];
            self.ranges = data.ranges || [];
            self.strings = data.strings || [];

            self.sheets = data.sheets || [];

            self.updateSheetMap();

            if (self.numOfDisplaySheets < 1) {
                for (var i = 0; i < 3; i++) self.addSheet();
            }

            return self;
        },

        //解析Sheet数据
        parseSheet: function (sheet) {
            var self = this;

            if (typeof sheet === "number") sheet = self.sheets[sheet];

            if (sheet._parsed) return self;
            sheet._parsed = true;

            var formats = self.formats,
                strings = self.strings,

                cells = [],

                dataCells = sheet.cells || [],
                count = dataCells.length,
                i = 0,

                dataCell,
                rowIndex,
                colIndex,

                maxCol = 0;

            for (; i < count; i++) {
                //dataCell => [row,col,type,value,format,style]
                dataCell = dataCells[i];
                rowIndex = dataCell[0];
                colIndex = dataCell[1];

                if (!cells[rowIndex]) cells[rowIndex] = [];
                cells[rowIndex][dataCell[1]] = { type: dataCell[2], value: strings[dataCell[3]], format: formats[dataCell[4]], style: dataCell[5] };

                if (colIndex > maxCol) maxCol = colIndex;
            }

            sheet.displayFormula = sheet.displayFormula !== false;
            sheet.displayGridlines = sheet.displayGridlines !== false;
            sheet.displayHeadings = sheet.displayHeadings !== false;

            sheet.colWidths = sheet.colWidths || {};
            sheet.rowHeights = sheet.rowHeights || {};
            sheet.rowStyles = sheet.rowStyles || {};
            sheet.colStyles = sheet.colStyles || {};
            sheet.hiddenRows = sheet.hiddenRows || [];
            sheet.hiddenCols = sheet.hiddenCols || [];
            sheet.merges = sheet.merges || [];
            sheet.links = sheet.links || [];
            sheet.comments = sheet.comments || [];
            sheet.richTexts = sheet.richTexts || [];
            sheet.shapes = sheet.shapes || [];

            sheet.cells = cells;

            sheet.maxRow = cells.length;
            sheet.maxCol = maxCol;

            return self;
        },

        //更新工作表映射
        updateSheetMap: function () {
            var self = this,
                sheets = [],
                mapSheet = {},
                numOfDisplaySheets = 0;

            self.sheets.forEach(function (sheet, i) {
                if (!sheet) return;

                sheet.index = i;
                sheets.push(sheet);
                mapSheet[sheet.name] = sheet;

                if (!sheet.isHidden) numOfDisplaySheets++;
            });

            self.sheets = sheets;
            self.mapSheet = mapSheet;

            self.numOfDisplaySheets = numOfDisplaySheets;

            return self;
        },

        //创建不重复的工作表名称
        _uniqueSheetName: function (str, start, s1, s2) {
            var i = 0, name;
            while (++i) {
                name = str + s1 + (start + i) + s2;
                if (!this.hasSheet(name)) return name;
            }
        },

        //创建不重复的工作表名称
        createSheetName: function () {
            return this._uniqueSheetName(Lang.DEF_SHEET_NAME, 0, "", "");
        },
        //克隆工作表名称 eg:Sheet1 => Sheet1 (2)
        cloneSheetName: function (name) {
            return this._uniqueSheetName(name, 1, " (", ")");
        },

        //是否是可视工作表
        isDisplaySheet: function (sheet) {
            return sheet && !sheet.isHidden;
        },
        //是否是可视工作表
        isDisplaySheetAt: function (index) {
            return this.isDisplaySheet(this.sheets[index]);
        },
        //是否是隐藏工作表
        isHiddenSheet: function (sheet) {
            return sheet && sheet.isHidden;
        },
        //是否是隐藏工作表
        isHiddenSheetAt: function (index) {
            return this.isHiddenSheet(this.sheets[index]);
        },
        //是否是空工作表
        isEmptySheet: function (index) {
            var sheet = this.sheets[index];
            if (!sheet) return true;

            var cells = sheet.cells;
            if (!cells || cells.length <= 0) return true;

            if (!sheet._parsed) return false;

            var len = cells.length, i = 0, row;
            for (; i < len; i++) {
                row = cells[i];
                if (row && row.length > 0) return false;
            }

            return true;
        },

        //是否存在指定名称的工作表
        hasSheet: function (name) {
            return this.mapSheet[name] != undefined;
        },
        //是否存在隐藏的工作表
        hasHiddenSheet: function () {
            return this.sheets.some(this.isHiddenSheet);
        },

        //插入工作表,可指定插入索引，若索引不存在，则添加到末尾
        insertSheet: function (sheet, index) {
            var self = this, sheets = self.sheets;

            if (sheet.index != undefined || self.hasSheet(sheet.name)) return Errors.SHEET_ALREADY_EXISTS;

            if (index == undefined || index == -1) {
                sheet.index = sheets.length;
                sheets.push(sheet);
                self.mapSheet[sheet.name] = sheet;
                if (!sheet.isHidden) self.numOfDisplaySheets++;
            } else {
                sheets.splice(index, 0, sheet);
                self.updateSheetMap();
            }

            return sheet;
        },
        //添加或新建工作表
        addSheet: function (sheet) {
            return this.insertSheet(sheet || { name: this.createSheetName() });
        },
        //克隆工作表
        cloneSheet: function (sheet) {
            var self = this;

            if (typeof sheet == "number") sheet = self.sheets[sheet];
            if (!sheet) return Errors.SHEET_NOT_FOUND;

            var newSheet = clone(sheet);
            newSheet.index = undefined;
            newSheet.name = self.cloneSheetName(sheet.name);
            return self.insertSheet(newSheet, sheet.index);
        },
        //隐藏工作表
        hideSheet: function (sheet) {
            var self = this;

            if (typeof sheet == "number") sheet = self.sheets[sheet];

            if (!sheet) return Errors.SHEET_NOT_FOUND;
            if (self.numOfDisplaySheets < 2) return Errors.SHEET_NUMBER_LESS_TWO;

            sheet.isHidden = true;
            self.index = self.getNextSheetIndex(sheet.index);
            self.numOfDisplaySheets--;

            return self;
        },
        //取消隐藏的工作表
        cancelHiddenSheets: function (indexs) {
            var self = this, sheets = self.sheets, lastSheetIndex;

            makeArray(indexs).forEach(function (index) {
                var sheet = sheets[+index];
                if (!sheet) return;

                sheet.isHidden = false;
                self.numOfDisplaySheets++;
                lastSheetIndex = sheet.index;
            });

            return lastSheetIndex;
        },
        //根据工作表名称获取工作表
        getSheet: function (name) {
            return this.mapSheet[name];
        },
        //根据索引获取工作表对象
        getSheetAt: function (index) {
            return this.sheets[index != undefined ? index : this.index];
        },
        //获取下一个可视工作表索引
        //includeSelf:是否包含自身
        getNextSheetIndex: function (index, includeSelf) {
            var self = this,
                sheets = self.sheets,
                len = sheets.length,
                start = includeSelf ? index : index + 1;

            //先向右查找
            while (start < len) {
                if (self.isDisplaySheetAt(start)) return start;
                start++;
            }

            //若右边无工作表,则向左查找
            while (--index >= 0) {
                if (self.isDisplaySheetAt(index)) return index;
            }
        },
        //获取下一个可视工作表
        getNextSheet: function (index, includeSelf) {
            var i = this.getNextSheetIndex(index, includeSelf);
            return this.sheets[i];
        },

        //获取隐藏的工作表
        getHiddenSheets: function () {
            return this.sheets.filter(this.isHiddenSheet);
        },

        //设置工作表
        setSheet: function (index, sheet) {
            var self = this;

            if (!sheet.name) return Errors.SHEET_NAME_NOT_FOUND;
            if (self.hasSheet(sheet.name)) return Errors.SHEET_ALREADY_EXISTS;

            self.sheets[index] = sheet;

            return self.updateSheetMap();
        },
        //重命名工作表
        renameSheet: function (sheet, name) {
            var self = this;

            if (typeof sheet == "number") sheet = self.sheets[sheet];

            if (!sheet) return Errors.SHEET_NOT_FOUND;
            if (!name) return Errors.PARAMETER_ERROR;
            if (name == sheet.name) return Errors.SHEET_NAME_IS_SAME;
            if (self.hasSheet(name)) return Errors.SHEET_NAME_ALREADY_EXISTS;

            delete self.mapSheet[sheet.name];
            sheet.name = name;
            self.mapSheet[name] = sheet;

            return self;
        },
        //移动工作表,返回移动的工作表或操作状态
        moveSheet: function (oldIndex, index) {
            var self = this;

            //工作薄内至少含有一张可视工作表
            if (self.numOfDisplaySheets < 2) return Errors.SHEET_NUMBER_LESS_TWO;

            var sheets = self.sheets, sheet = sheets[oldIndex], len = sheets.length;

            if (!sheet) return Errors.SHEET_NOT_FOUND;
            if (oldIndex < 0 || oldIndex >= len) return Errors.PARAMETER_ERROR;

            if (index > len || index == -1) index = len - 1;
            else if (index < 0) index = 0;

            if (index == oldIndex) return Errors.SHEET_INDEX_IS_SAME;

            sheets.del(oldIndex);
            if (index == len - 1) sheets.push(sheet);
            else sheets.splice(index, 0, sheet);

            self.updateSheetMap();

            return sheet;
        },
        //移除工作表,返回移除的工作表或操作状态
        removeSheet: function (index) {
            var self = this;

            //工作薄内至少含有一张可视工作表
            if (self.numOfDisplaySheets < 2) return Errors.SHEET_NUMBER_LESS_TWO;

            var sheets = self.sheets, lastIndex = sheets.length - 1;
            if (index == -1) index = lastIndex;
            if (index == undefined || index < 0 || index > lastIndex) return Errors.PARAMETER_ERROR;

            var sheet = sheets[index];
            if (!sheet) return Errors.SHEET_NOT_FOUND;

            sheets.del(index);
            if (index != lastIndex) {
                self.updateSheetMap();
            } else {
                delete sheet.index;
                delete self.mapSheet[sheet.name];
                self.numOfDisplaySheets--;
            }

            return sheet;
        },

        getCellFont: function (index) {
            return this.fonts[index];
        },

        getCellStyle: function (index) {
            return this.styles[index];
        },

        getCellRange: function (index) {
            return this.ranges[index];
        },

        getData: function () {
            var self = this,
                data = {},

                formats = [],
                strings = [],

                mapFormat = {},
                mapString = {},

                dataSheets = [];

            ["index", "title", "subject", "author", "company", "fonts", "styles", "ranges"].forEach(function (prop) {
                data[prop] = self[prop];
            });

            self.sheets.forEach(function (sheet) {
                var dataSheet = extend({}, sheet);

                delete dataSheet["index"];

                if (!dataSheet._parsed) {
                    dataSheets.push(dataSheet);
                    return;
                }

                var dataCells = [];
                dataSheet.cells.forEach(function (row, i) {
                    if (!row) return;

                    row.forEach(function (cell, j) {
                        if (!cell) return;

                        var format = cell.format,
                            value = cell.value + "",

                            i_format = format ? mapFormat[format] : 0,
                            i_string = value ? mapString[value] : 0;

                        if (i_format == undefined) {
                            i_format = mapFormat[format] = formats.length;
                            formats.push(format);
                        }

                        if (i_string == undefined) {
                            i_string = mapString[value] = strings.length;
                            strings.push(value);
                        }

                        dataCells.push([i, j, cell.type, i_string, i_format, cell.style]);
                    });
                });

                dataSheet.cells = dataCells;

                dataSheets.push(dataSheet);
            });

            data.formats = formats;
            data.strings = strings;
            data.sheets = dataSheets;

            return data;
        }
    });

    SS.Workbook = Workbook;

})(window);