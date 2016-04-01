/// <reference path="ui.js" />
/*
* spreadsheets.js
* author:devin87@qq.com
* update: 2016/04/01 17:15
*/
(function (window, undefined) {
    "use strict";

    var async = Q.async,
        factory = Q.factory,

        Listener = Q.Listener,

        SS = Q.SS,
        Errors = SS.Errors,
        Events = SS.Events,
        Workbook = SS.Workbook,
        UI = SS.UI,

        Lang = Q.SSLang;

    //Excel表格对象
    function Spreadsheets(ops) {
        var self = this;

        //自定义事件监听器
        self._listener = new Listener(Object.values(Events), self);

        self.ops = ops;
        self.workbook = new Workbook;
        self.ui = new UI(self);

        self.init();
    }

    factory(Spreadsheets).extend({
        //初始化
        init: function () {
            var self = this,
                ops = self.ops;

            self.load(ops.data);
        },
        //添加自定义事件
        on: function (type, fn) {
            this._listener.add(type, fn);
            return this;
        },
        //移除自定义事件
        off: function (type, fn) {
            this._listener.remove(type, fn);
            return this;
        },
        //触发自定义事件
        trigger: function (type, args) {
            this._listener.trigger(type, args);
            return this;
        },
        //加载workbook数据
        load: function (data) {
            var self = this,
                workbook = self.workbook;

            workbook.parse(data);
            self.trigger(Events.WORKBOOK_READY, workbook);

            return self.loadSheetAt(workbook.index);
        },

        //加载工作表
        loadSheet: function (sheet) {
            var self = this,
                workbook = self.workbook;

            workbook.parseSheet(sheet);
            self.trigger(Events.SHEET_READY, sheet);

            self.ui.loadSheet(sheet);

            workbook.index = sheet.index;
            self.sheet = sheet;

            return self.trigger(Events.SHEET_LOAD, sheet);
        },
        //加载工作表
        loadSheetAt: function (index) {
            var self = this,
                workbook = self.workbook;

            index = +index || 0;
            var sheet = workbook.getSheetAt(index);
            if (!workbook.isDisplaySheet(sheet)) sheet = workbook.getNextSheet(index);

            return self.loadSheet(sheet);
        },
        //添加工作表
        addSheet: function () {
            var self = this,
                sheet = self.workbook.addSheet();

            self.ui.addSheet(sheet);

            return self;
        },
        //克隆工作表
        cloneSheet: function (index) {
            var self = this,
                sheet = self.workbook.cloneSheet(index);

            self.ui.cloneSheet(sheet);

            return self;
        },
        //隐藏工作表
        hideSheet: function (index) {
            var self = this,
                workbook = self.workbook,
                sheet = workbook.getSheetAt(index),
                title = Lang.HIDDEN_SHEET + " - " + sheet.name,

                result = workbook.hideSheet(index),
                errMsg;

            //删除失败
            if (typeof result == "number") {
                switch (result) {
                    case Errors.SHEET_NUMBER_LESS_TWO: errMsg = Lang.SHEET_NUMBER_LESS_TWO; break;
                    case Errors.PARAMETER_ERROR: errMsg = Lang.PARAMETER_ERROR; break;
                    case Errors.SHEET_NOT_FOUND: errMsg = Lang.HIDDEN_SHEET_NOT_FOUND; break;
                }

                return alertError(errMsg, { title: title });
            } else {
                self.ui.cloneSheet(sheet);
            }

            return self;
        },

        //取消隐藏的工作表
        cancelHiddenSheets: function (indexs) {
            var self = this,
                index = self.workbook.cancelHiddenSheets(indexs);

            if (index != undefined) {
                self.ui.cancelHiddenSheets(index);
            }

            return self;
        },
        //移动工作表
        moveSheet: function (oldIndex, index) {
            var self = this,
                workbook = self.workbook,
                result = workbook.moveSheet(oldIndex, index);

            if (typeof result != "number") {
                self.ui.moveSheet(result);
            }

            return self;
        },
        //移除工作表
        removeSheet: function (index) {
            var self = this,
                workbook = self.workbook,
                sheet = workbook.getSheetAt(index),
                title = Lang.REMOVE_SHEET + " - " + sheet.name;

            var removeSheet = function (yes) {
                if (!yes) return;

                var result = workbook.removeSheet(index), errMsg;

                //删除失败
                if (typeof result == "number") {
                    switch (result) {
                        case Errors.SHEET_NUMBER_LESS_TWO: errMsg = Lang.SHEET_NUMBER_LESS_TWO; break;
                        case Errors.PARAMETER_ERROR: errMsg = Lang.PARAMETER_ERROR; break;
                        case Errors.SHEET_NOT_FOUND: errMsg = Lang.REMOVE_SHEET_NOT_FOUND; break;
                    }

                    return alertError(errMsg, { title: title });
                }

                self.ui.removeSheet(result);
            };

            //如果工作表无数据,则无需用户确认,直接删除
            if (workbook.isEmptySheet(index)) return removeSheet(true);

            Q.confirm(Lang.REMOVE_SHEET_CONFIRM, removeSheet, { title: title });

            return self;
        },
        //重命名工作表
        renameSheet: function (index) {
            var self = this,
                workbook = self.workbook,
                sheet = workbook.getSheetAt(index),
                title = Lang.RENAME_SHEET + " - " + sheet.name;

            Q.prompt(Lang.INPUT_SHEET_NAME, function (value) {
                if (!value) return false;

                var result = workbook.renameSheet(sheet, value);

                //删除失败
                if (typeof result == "number") {
                    switch (result) {
                        case Errors.SHEET_NAME_IS_SAME: errMsg = Lang.SHEET_NAME_IS_SAME; break;
                        case Errors.SHEET_NAME_ALREADY_EXISTS: errMsg = Lang.SHEET_NAME_ALREADY_EXISTS; break;
                        case Errors.PARAMETER_ERROR: errMsg = Lang.PARAMETER_ERROR; break;
                        case Errors.SHEET_NOT_FOUND: errMsg = Lang.RENAME_SHEET_NOT_FOUND; break;
                    }

                    return alertError(errMsg, { title: title });
                }

                self.ui.renameSheet(sheet, value);

            }, { title: title });

            return self;
        },

        scrollToRow: function (row, isUpdateScrollbar) {
            return self.ui.scrollToRow(row, isUpdateScrollbar);
        },
        scrollToCol: function (col, isUpdateScrollbar) {
            return self.ui.scrollToCol(col, isUpdateScrollbar);
        }
    });

    Q.Spreadsheets = Spreadsheets;

})(window);