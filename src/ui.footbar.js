/// <reference path="ui.js" />
/*
* ui.footbar.js 底部栏操作
* author:devin87@qq.com
* update: 2016/03/30 15:48
*/
(function (window, undefined) {
    "use strict";

    var extend = Q.extend,
        vals = Q.vals,
        toMap = Q.toMap,

        createEle = Q.createEle,

        measureText = Q.getMeasureText(),
        view = Q.view,

        ContextMenu = Q.ContextMenu,

        SS = Q.SS,
        ERROR = SS.ERROR,
        UI = SS.UI,
        Lang = Q.SSLang,

        picker = UI.picker;

    //tab 右键菜单类型
    var TAB_MENU_INSERT = 1,
        TAB_MENU_REMOVE = 2,
        TAB_MENU_RENAME = 3,
        TAB_MENU_CLONE = 4,
        TAB_MENU_MOVE_TO_LEFT = 5,
        TAB_MENU_MOVE_TO_RIGHT = 6,
        TAB_MENU_MOVE_TO_TOP = 7,
        TAB_MENU_MOVE_TO_BOTTOM = 8,
        TAB_MENU_HIDDEN = 9,
        TAB_MENU_CANCEL_HIDDEN = 10;

    //tab 右键菜单数据
    var MENU_TAB_DATA = Q.processMenuWidth({
        width: 140,
        items: [
            { id: TAB_MENU_INSERT, text: Lang.INSERT },
            { id: TAB_MENU_REMOVE, text: Lang.REMOVE },
            { id: TAB_MENU_RENAME, text: Lang.RENAME },
            { split: true },
            { id: TAB_MENU_CLONE, text: Lang.CLONE },
            {
                text: Lang.MOVE,
                group: {
                    items: [
                        { id: TAB_MENU_MOVE_TO_LEFT, text: Lang.MOVE_TO_LEFT },
                        { id: TAB_MENU_MOVE_TO_RIGHT, text: Lang.MOVE_TO_RIGHT },
                        { id: TAB_MENU_MOVE_TO_TOP, text: Lang.MOVE_TO_TOP },
                        { id: TAB_MENU_MOVE_TO_BOTTOM, text: Lang.MOVE_TO_BOTTOM }
                    ]
                }
            },
            //链接菜单:颜色选择器
            { text: Lang.TAB_COLOR, group: picker },
            { split: true },
            { id: TAB_MENU_HIDDEN, text: Lang.HIDDEN },
            { id: TAB_MENU_CANCEL_HIDDEN, text: Lang.CANCEL_HIDDEN }
        ]
    });

    UI.extend({
        //初始化底部栏
        initFootbar: function () {
            var self = this,
                workbook = self.workbook;

            var html =
                '<table class="sst-table">' +
                    '<tr>' +
                        '<td class="sst-w1"><div class="sst-arrow-up sst-view-all" title="' + Lang.VIEW_ALL_SHEETS + '"></div></td>' +
                        '<td class="sst-w2"><div class="sst-tabs"></div></td>' +
                        '<td class="sst-w3"><div class="sst-resize"></div></td>' +
                        '<td class="sst-w4"><div class="ss-hscroll"><div class="ss-total-width"></div></div></td>' +
                        '<td class="sst-w5"></td>' +
                    '</tr>' +
                '</table>';

            var elFootbar = $(".ss-footbar", self.elView).html(html)[0],
                elTabs = $(".sst-tabs", elFootbar)[0],
                elResize = $(".sst-resize", elFootbar)[0],
                elHScroll = $(".ss-hscroll", elFootbar)[0];

            self.elFootbar = elFootbar;
            self.elTabs = elTabs;
            self.elHScroll = elHScroll;

            //水平滚动条宽度
            self._HScrollWidth = elHScroll.offsetWidth;
            self._HScrollHeight = elHScroll.offsetHeight;

            //绘制Sheet标签
            self.drawTabs();

            //显示查看所有工作表菜单
            $(".sst-view-all", elFootbar).click(function (e) {
                var menu = self.menuViewAll || self.createMenuViewAll();

                menu.set({ rangeX: 0, rangeY: elFootbar.offsetTop }).toggle(this.offsetLeft + 2, elFootbar.offsetTop);

                return false;
            });

            var tab_sheet_index;

            //底部标签栏事件
            $(elTabs).on("click", ".sst-item", function () {
                self.runApi("loadSheetAt", this.x);
            }).on("contextmenu", ".sst-item", function (e) {
                var el = this,
                    elName = $(".sst-name", el)[0],
                    menuTab = self.menuTab;

                tab_sheet_index = el.x;

                if (!menuTab) {
                    self.menuTab = menuTab = new ContextMenu(MENU_TAB_DATA);

                    menuTab.onclick = function (e, item) {
                        switch (item.data.id) {
                            case TAB_MENU_INSERT: self.runApi("addSheet"); break;
                            case TAB_MENU_REMOVE: self.runApi("removeSheet", tab_sheet_index); break;
                            case TAB_MENU_RENAME: self.runApi("renameSheet", tab_sheet_index); break;
                            case TAB_MENU_CLONE: self.runApi("cloneSheet", tab_sheet_index); break;
                            case TAB_MENU_MOVE_TO_LEFT: self.runApi("moveSheet", tab_sheet_index, tab_sheet_index - 1); break;
                            case TAB_MENU_MOVE_TO_RIGHT: self.runApi("moveSheet", tab_sheet_index, tab_sheet_index + 1); break;
                            case TAB_MENU_MOVE_TO_TOP: self.runApi("moveSheet", tab_sheet_index, 0); break;
                            case TAB_MENU_MOVE_TO_BOTTOM: self.runApi("moveSheet", tab_sheet_index, -1); break;
                            case TAB_MENU_HIDDEN: self.runApi("hideSheet", tab_sheet_index); break;
                            case TAB_MENU_CANCEL_HIDDEN: self.showHiddenSheets(); break;
                        }
                    };
                }

                menuTab[workbook.hasHiddenSheet() ? "enableItems" : "disableItems"](TAB_MENU_CANCEL_HIDDEN);

                menuTab.set({ rangeX: 0, rangeY: elFootbar.offsetTop + 5 }).show(e.x, e.y);

                picker.set({
                    color: elName.style.color,
                    callback: function (color) {
                        var sheet = workbook.getSheetAt(el.x);
                        elName.style.color = sheet.tabColor = color;

                        menuTab.hide();
                    }
                });

                self.runApi("loadSheetAt", el.x);

                return false;
            }).on("click", ".sst-insert", function () {
                //新建工作表
                self.runApi("addSheet");
            });

            //底部水平滚动条滚动事件
            $(elHScroll).on("scroll", function (e) {
                if (!self.loaded) return;

                if (self.skipScrollEvent) {
                    self.skipScrollEvent = false;
                    return;
                }

                var left = this.scrollLeft,
                    col = self.findCol(left);

                self.scrollToColAsync(col, false);
            });
        },
        //创建“查看所有工作表”菜单,若菜单已存在,则先销毁再创建
        createMenuViewAll: function () {
            var self = this,
                workbook = self.workbook,

                menuItems = [],
                maxTextWidth = 70,
                menu = self.menuViewAll;

            if (menu) menu.destroy();

            workbook.sheets.forEach(function (sheet) {
                if (!workbook.isDisplaySheet(sheet)) return;

                var index = sheet.index,
                    name = sheet.name,
                    //tabColor = sheet.tabColor,
                    textWidth = measureText.getWidth(name);

                if (textWidth > maxTextWidth) maxTextWidth = textWidth;

                menuItems.push({ id: index, text: name, x: index });
            });

            self.menuViewAll = menu = new ContextMenu({
                width: 50 + maxTextWidth,
                className: "checked-panel",
                items: menuItems
            }, { fixedWidth: true });

            menu.onclick = function () {
                self.runApi("loadSheetAt", this.x);
            };

            var item = menu.getItem(workbook.index);
            if (item) $(item.node).addClass("ss-active");

            return menu;
        },
        //绘制sheet选项卡
        drawTabs: function () {
            var self = this,
                workbook = self.workbook,
                sheets = workbook.sheets,
                sheetIndex = workbook.index,

                elTabs = self.elTabs;

            var html =
                sheets.map(function (sheet) {
                    if (!workbook.isDisplaySheet(sheet)) return '';

                    var index = sheet.index,
                        name = sheet.name,
                        tabColor = sheet.tabColor;

                    var html =
                        '<div class="sst-item' + (index == sheetIndex ? ' sst-on' : '') + '" title="' + name + '"' + (tabColor ? ' style="color:' + tabColor + ';"' : '') + '>' +
                            '<div class="sst-name">' + name + '</div>' +
                        '</div>';

                    return html;
                }).join('') +
                '<div class="sst-insert" title="' + Lang.INSERT_SHEET + '">' +
                    '<div class="sst-name">+</div>' +
                '</div>';

            $(elTabs).html(html);

            var mapTab = {};
            $(".sst-item", elTabs).each(function (i, tab) {
                var sheet = workbook.getSheet(tab.title);
                tab.x = sheet.index;

                mapTab[tab.x] = tab;
            });

            self.mapTab = mapTab;

            var menu = self.menuViewAll;
            if (menu) {
                menu.destroy();
                self.menuViewAll = undefined;
            }
        },
        //激活sheet选项卡
        activeTab: function (sheet) {
            var self = this,
                workbook = self.workbook,
                mapTab = self.mapTab,
                menu = self.menuViewAll;

            $(mapTab[workbook.index]).removeClass("sst-on");
            $(mapTab[sheet.index]).addClass("sst-on");

            if (menu) {
                var lastItem = menu.getItem(workbook.index),
                    item = menu.getItem(sheet.index);

                if (lastItem) $(lastItem.node).removeClass("ss-active");
                if (item) $(item.node).addClass("ss-active");
            }
        },
        //重命名sheet选项卡
        renameTab: function (sheet, name) {
            var self = this,
                workbook = self.workbook,
                elTab = self.mapTab[sheet.index];

            $(".sst-name", elTab).html(name);

            self.createMenuViewAll();

            return self;
        },
        //显示隐藏的工作表
        showHiddenSheets: function () {
            var self = this,
                workbook = self.workbook;

            var html =
                '<div class="text">' + Lang.CANCEL_HIDDEN_SHEETS + '</div>' +
                '<select class="sel-hidden-sheets" multiple="multiple">' +
                    workbook.sheets.map(function (sheet) {
                        return workbook.isHiddenSheet(sheet) ? '<option value="' + sheet.index + '">' + sheet.name + '</option>' : '';
                    }).join('') +
                '</select>';

            var box = new Q.WinBox({ title: Lang.CANCEL_HIDDEN, html: html + Q.bottom(2), width: 360, close: "remove", mask: true });

            box.on(".x-submit", "click", function () {
                var sel = box.get(".sel-hidden-sheets"),
                    indexs = vals(sel.selectedOptions, "value");

                if (indexs.length > 0) self.runApi("cancelHiddenSheets", indexs);

                box.remove();
            }).on(".x-cancel", "click", "remove");
        }
    });

    UI.register("initFootbar");

})(window);