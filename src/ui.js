/// <reference path="workbook.js" />
/// <reference path="events.js" />
/*
* ui.js
* update: 2015/11/23 17:33
*/
(function (window, undefined) {
    "use strict";

    var slice = Array.prototype.slice,

        isFunc = Q.isFunc,
        extend = Q.extend,
        toMap = Q.toMap,

        alertError = Q.alertError,

        factory = Q.factory,
        view = Q.view,

        Listener = Q.Listener,
        ColorPicker = Q.ColorPicker,

        SS = Q.SS,
        Errors = SS.Errors,
        Events = SS.Events,
        Workbook = SS.Workbook,
        Lang = Q.SSLang;

    function UI(api) {
        var self = this;

        self.api = api;
        self.ops = api.ops;
        self.workbook = api.workbook;

        self.init();
    }

    factory(UI).extend({
        init: function () {
            var self = this,
                ops = self.ops,
                height = ops.height || "auto",
                elView = ops.view;

            ops.height = height;
            self.elView = elView;

            var html =
                '<div class="ss-toolbar"></div>' +
                '<div class="ss-formulabar"></div>' +
                '<div class="ss-sheet"></div>' +
                '<div class="ss-footbar"></div>';

            $(elView).addClass("ss-view").height(height == "auto" ? view.getHeight() : height).html(html);

            if (UI.init) UI.init(self);

            $(window).on("resize", function () {
                self.updateSize();
            });
        },
        trigger: function (type, args) {
            this.api.trigger(type, args);
            return this;
        },
        updateSize: function () {
            var self = this;

            if (self.ops.height == "auto") $(self.elView).height(view.getHeight());
            self.run("resizeGrid");

            return self.trigger(Events.UI_RESIZE);
        },
        run: function (action) {
            var self = this, fn = self[action];
            if (isFunc(fn)) fn.apply(self, slice.call(arguments, 1));

            return self;
        },
        runApi: function (action) {
            var self = this, api = self.api || {}, fn = api[action];
            if (isFunc(fn)) fn.apply(api, slice.call(arguments, 1));

            return self;
        },
        loadSheet: function (sheet) {
            var self = this;
            self.sheet = sheet;

            return self.run("activeTab", sheet).run("drawSheet", sheet);
        },
        draw: function (sheet) {
            var self = this,

                workbook = self.workbook,
                fonts = workbook.fonts,
                styles = workbook.styles,

                colWidths = sheet.colWidths,
                rowHeights = sheet.rowHeights,

                rowStyles = sheet.rowStyles,
                colStyles = sheet.colStyles,

                hiddenRows = sheet.hiddenRows,
                hiddenCols = sheet.hiddenCols,

                merges = sheet.merges,
                links = sheet.links,
                comments = sheet.comments,
                richTexts = sheet.richTexts,

                cells = sheet.cells,

                colWidth = sheet.colWidth || 80,
                rowHeight = sheet.rowHeight || 20,

                style = sheet.style != undefined ? styles[sheet.style] : undefined;
        },

        addSheet: function () {
            return this.run("drawTabs");
        },
        cloneSheet: function (index) {
            return this.run("drawTabs");
        },
        hideSheet: function (index) {
            return this.run("drawTabs");
        },
        cancelHiddenSheets: function (indexs) {
            return this.run("drawTabs");
        },
        moveSheet: function (oldIndex, index) {
            return this.run("drawTabs");
        },
        removeSheet: function (index) {
            return this.run("drawTabs");
        },
        renameSheet: function (sheet, name) {
            return this.run("renameTab", sheet, name);
        }
    });

    extend(UI, {
        modules: [],
        init: function (scope) {
            UI.scope = scope;
            UI.loaded = true;

            UI.modules.forEach(UI.fire);
        },
        fire: function (module) {
            if (isFunc(module)) module(UI.scope);
            else if (typeof module == "string") UI.scope[module]();
        },
        register: function (module) {
            UI.modules.push(module);

            if (UI.loaded) UI.fire(module);
        },

        picker: new ColorPicker({ isHex: true })
    });

    SS.UI = UI;

})(window);