/// <reference path="ui.js" />
/*
* ui.grid.js 表格操作
* update: 2016/04/12 11:58
*/
(function (window, undefined) {
    "use strict";

    var extend = Q.extend,
        async = Q.async,
        fire = Q.fire,

        toMap = Q.toMap,

        getFirst = Q.getFirst,
        getNext = Q.getNext,
        getLast = Q.getLast,
        createEle = Q.createEle,

        getOffset = Q.offset,

        findInList = Q.findInList,

        measureText = Q.getMeasureText(),
        view = Q.view,

        ContextMenu = Q.ContextMenu,

        SS = Q.SS,
        ERROR = SS.ERROR,
        Lang = SS.Lang,
        UI = SS.UI,

        getColName = SS.getColName,

        picker = UI.picker;

    var DEF_COL_WIDTH = 80,
        DEF_ROW_HEIGHT = 20,

        DEF_ROW_NUM_WIDTH = 45,
        DEF_COL_HEAD_HEIGHT = 22,

        INFINITY_VALUE = -1,

        RANGE_NO_MERGE = -1,           //非合并区域
        RANGE_IS_MERGE = 0,            //合并区域
        RANGE_IS_MERGE_FIRST = 1,      //合并区域中的第一个单元格
        RANGE_IS_MERGE_OTHER = 2;      //合并区域中的其它单元格

    function log(msg) {
        //if (window.console) console.log.apply(console, arguments);
        //Q.alert(msg);
    }

    //计算单元格偏移  eg:{ 1:67,  3:45 }
    //indexs:  非默认宽高的行或列索引列表   => [1,3]
    //offsets: 非默认宽高的宽高累加偏移列表 => [67,112]
    //v: 默认的宽或高
    //i: 行或列索引
    function calcCellOffset(indexs, offsets, v, i) {
        var index = findInList(indexs, i), len = indexs.length;
        //-2表示小于列表第一项
        if (index == -2 || i == indexs[0]) return i * v;
        //-1表示大于列表最后一项
        if (index == -1) return offsets[len - 1] + (i - len) * v;

        if (i == indexs[index]) index--;
        return offsets[index] + (i - index - 1) * v;
    }

    //根据偏移查找单元格行或列索引 eg:{ 1:67,  3:45 }
    //offsets: 非默认宽高的宽高累加偏移列表 => [20,87,152]
    //indexs:  非默认宽高的行或列索引列表   => [1,2,4]
    //v: 默认的宽或高
    //n: 偏移值
    function findCellIndex(offsets, indexs, v, n) {
        //非默认宽高的数量
        var len = offsets.length;
        if (len == 0 || n < offsets[0]) return Math.floor(n / v);

        var last = offsets[--len];
        if (n >= last) return indexs[len] + Math.floor((n - last) / v);

        var i = findInList(offsets, n),
            index = indexs[i],
            next = indexs[i + 1],
            fix = n - offsets[i];

        if (fix == 0) return index;

        index += Math.floor(fix / v);
        return index < next ? index : next - 1;
    }

    //计算区域(支持全选行或全选列 eg:[15,INFINITY_VALUE,16,INFINITY_VALUE],[INFINITY_VALUE,3,INFINITY_VALUE,5])
    function computeRange(list, firstRow, firstCol, lastRow, lastCol) {
        var _firstRow = firstRow,
            _firstCol = firstCol,
            _lastRow = lastRow,
            _lastCol = lastCol,

            skipRow = _firstRow == INFINITY_VALUE,
            skipCol = _firstCol == INFINITY_VALUE,

            data;

        for (var i = 0, len = list.length; i < len; i++) {
            if (skipRow && skipCol) break;

            data = list[i];
            if (skipRow || data[0] == INFINITY_VALUE) {  //全选行
                if (!skipRow) {
                    skipRow = true;
                    _firstRow = _lastRow = INFINITY_VALUE;
                }
                if (skipCol || data[1] == INFINITY_VALUE) {
                    _firstCol = _lastCol = INFINITY_VALUE;
                    break;
                }

                if (Math.abs(data[1] - firstCol + data[3] - lastCol) <= lastCol - firstCol + data[3] - data[1]) {
                    if (data[1] < _firstCol) _firstCol = data[1];
                    if (data[3] > _lastCol) _lastCol = data[3];
                }
            } else if (skipCol || data[1] == INFINITY_VALUE) {  //全选列
                if (!skipCol) {
                    skipCol = true;
                    _firstCol = _lastCol = INFINITY_VALUE;
                }
                if (skipRow || data[0] == INFINITY_VALUE) {
                    _firstRow = _lastRow = INFINITY_VALUE;
                    break;
                }

                if (Math.abs(data[0] - firstRow + data[2] - lastRow) <= lastRow - firstRow + data[2] - data[0]) {
                    if (data[0] < _firstRow) _firstRow = data[0];
                    if (data[2] > _lastRow) _lastRow = data[2];
                }
            } else {
                //判断矩形相交
                if (Math.abs(data[0] - firstRow + data[2] - lastRow) <= lastRow - firstRow + data[2] - data[0] && Math.abs(data[1] - firstCol + data[3] - lastCol) <= lastCol - firstCol + data[3] - data[1]) {
                    if (data[0] < _firstRow) _firstRow = data[0];
                    if (data[1] < _firstCol) _firstCol = data[1];
                    if (data[2] > _lastRow) _lastRow = data[2];
                    if (data[3] > _lastCol) _lastCol = data[3];
                }
            }

            if (_firstRow != firstRow || _firstCol != firstCol || _lastRow != lastRow || _lastCol != lastCol) return computeRange(list, _firstRow, _firstCol, _lastRow, _lastCol);
        }

        return [_firstRow, _firstCol, _lastRow, _lastCol];
    }

    UI.extend({
        initGrid: function () {
            var self = this,
                workbook = self.workbook;

            var html =
                '<div class="ss-grid ssg-1">' +
                    '<div class="ss-cells"></div>' +
                '</div>' +
                '<div class="ss-grid ssg-2">' +
                    '<div class="ss-cells"></div>' +
                '</div>' +
                '<div class="ss-grid ssg-3">' +
                    '<div class="ss-cells"></div>' +
                '</div>' +
                '<div class="ss-grid ssg-4">' +
                    '<div class="ss-cells"></div>' +
                '</div>' +
                '<div class="ss-mark"></div>' +
                '<div class="ss-vscroll">' +
                    '<div class="ss-total-height"></div>' +
                '</div>';

            var elSheet = $(".ss-sheet", self.elView).html(html)[0];

            self.elSheet = elSheet;
            self.elGrid1 = getFirst(elSheet);
            self.elGrid2 = getNext(self.elGrid1);
            self.elGrid3 = getNext(self.elGrid2);
            self.elGrid4 = getNext(self.elGrid3);

            self.elGridCells1 = getFirst(self.elGrid1);
            self.elGridCells2 = getFirst(self.elGrid2);
            self.elGridCells3 = getFirst(self.elGrid3);
            self.elGridCells4 = getFirst(self.elGrid4);

            self.elMark = getNext(self.elGrid4);
            self.elVScroll = getLast(elSheet);
            self._VScrollbarWidth = self.elVScroll.offsetWidth || 17;

            self._rowNumWidth = DEF_ROW_NUM_WIDTH;
            self._colHeadHeight = DEF_COL_HEAD_HEIGHT;

            return self.initGridSize().initGridEvent();
        },

        initGridSize: function () {
            var self = this,

                elView = self.elView,
                elSheet = self.elSheet,
                elMark = self.elMark,
                elFootbar = self.elFootbar,

                height_other = elSheet.offsetTop - elView.offsetTop + (elFootbar ? elFootbar.offsetHeight : 0),

                GRID_WIDTH = elView.clientWidth,
                GRID_HEIGHT = elView.clientHeight - height_other;

            self.GRID_WIDTH = GRID_WIDTH;
            self.GRID_HEIGHT = GRID_HEIGHT;

            elSheet.style.cssText = 'width:' + GRID_WIDTH + 'px;height:' + GRID_HEIGHT + 'px;';
            elMark.style.cssText = 'width:' + (GRID_WIDTH - self._VScrollbarWidth) + 'px;height:' + GRID_HEIGHT + 'px;';

            return self;
        },

        initGridEvent: function () {
            var self = this,

                elSheet = self.elSheet,
                elMark = self.elMark,
                elVScroll = self.elVScroll;

            //垂直滚动条滚动事件
            $(elVScroll).on("scroll", function (e) {
                if (!self.loaded) return;

                if (self.skipScrollEvent) {
                    self.skipScrollEvent = false;
                    return;
                }

                var top = this.scrollTop,
                    row = self.findRow(top);

                self.scrollToRowAsync(row, false);
            });

            var firstRow,
                firstCol,
                lastRow,
                lastCol,

                lineRow,
                lineCol,

                is_mousedown,
                is_resize_row,
                is_resize_col,
                is_rightclick;

            //根据鼠标坐标获取单元格索引
            var get_cell_index = function (e) {
                var offset = getOffset(elSheet),
                    left = e.clientX - offset.left,
                    top = e.clientY - offset.top;

                if (left > self.FIXED_WIDTH) left += self._scrollLeft;
                if (top > self.FIXED_HEIGHT) top += self._scrollTop;

                var row = self.findRow(top),
                    col = self.findCol(left);

                lineRow = lineCol = 0;

                //判断是否正好处于行或列的起始位置
                if (row == INFINITY_VALUE) {
                    var colLeft = self.getColLeft(col);
                    if (left == colLeft) lineCol = col;
                    else if (left > colLeft) {
                        if (left < colLeft + 5) lineCol = col;
                        else if (left > self.getColLeft(col + 1) - 5) lineCol = col + 1;
                    }
                } else if (col == INFINITY_VALUE) {
                    var rowTop = self.getRowTop(row);
                    if (top == rowTop) lineRow = row;
                    else if (top > rowTop) {
                        if (top < rowTop + 5) lineRow = row;
                        else if (top > self.getRowTop(row + 1) - 5) lineRow = row + 1;
                    }
                }

                return { row: row, col: col };
            };

            $(elSheet).on({
                //鼠标中键滚轮事件
                wheel: function (e) {
                    //e.delta: 向上滚动为1,向下滚动为-1
                    var scrollRow = self._scrollRow,
                        row = scrollRow - 3 * e.delta;  //每次滚动3行

                    self.scrollToRowAsync(row, true);
                },

                mousedown: function (e) {
                    var target = e.target;
                    if (target == elVScroll) return;

                    is_mousedown = true;
                    is_rightclick = e.which == 3 || e.button == 2;   //是否右键菜单

                    var ci = get_cell_index(e);
                    lastRow = firstRow = ci.row;
                    lastCol = firstCol = ci.col;

                    if (lineRow) {
                        is_resize_row = true;
                        self.showDragHLine(lineRow);
                    } else if (lineCol) {
                        is_resize_col = true;
                        self.showDragVLine(lineCol);
                    } else {
                        self.selectCells(firstRow, firstCol);
                    }
                },

                mousemove: function (e) {
                    var ci = get_cell_index(e),
                        row = ci.row,
                        col = ci.col,
                        cursor = "cell";

                    if (row == INFINITY_VALUE || col == INFINITY_VALUE) {
                        cursor = "default";
                        if (lineRow) cursor = "n-resize";
                        else if (lineCol) cursor = "e-resize";
                    }

                    elMark.style.cursor = cursor;

                    if (!is_mousedown || (row == lastRow && col == lastCol)) return;

                    lastRow = row;
                    lastCol = col;

                    self.selectCells(firstRow, firstCol, lastRow, lastCol);
                },

                mouseup: function (e) {
                    is_mousedown = false;

                    if (is_resize_row) {
                        is_resize_row = false;
                        self.hideDragHLine();
                    } else if (is_resize_col) {
                        is_resize_col = false;
                        self.hideDragVLine();
                    }

                    var target = e.target;
                },

                contextmenu: function (e) {
                    return false;
                }
            });

            return self;
        },

        resizeGrid: function () {
            return this.initGridSize().drawSheet();
        },

        updateColsOffset: function () {
            var self = this,
                sheet = self.sheet,

                map_col_width = sheet.colWidths || {},

                list_col_width_index = [],
                list_col_width_offset = [],

                list_col_left = [],
                list_col_left_index = [],

                offset_width = 0,

                list_col_offset = [self._rowNumWidth],
                def_col_width = self._colWidth,
                i;

            var d1 = Date.now();

            Object.keys(map_col_width).sort(function (a, b) {
                return a - b;
            }).forEach(function (key, i) {
                var col = +key;
                list_col_width_index.push(col);

                offset_width += map_col_width[key];
                list_col_width_offset.push(offset_width);

                if (i == 0) {
                    list_col_left.push(col * def_col_width);
                    list_col_left_index.push(col);
                }
                list_col_left.push(offset_width + (col - i) * def_col_width);
                list_col_left_index.push(col + 1);
            });

            var d2 = Date.now();
            log("updateColsOffset: " + (d2 - d1) + "ms");

            self._colWidthIndexs = list_col_width_index;
            self._colWidthOffsets = list_col_width_offset;
            self._colOffsets = list_col_offset;

            self._colLefts = list_col_left;
            self._colLeftIndexs = list_col_left_index;

            //预缓存数据
            //for (i = 1; i < 200; i++) {
            //    list_col_offset[i] = (map_col_width[i - 1] || def_col_width) + list_col_offset[i - 1];
            //}

            return self;
        },
        updateRowsOffset: function () {
            var self = this,
                sheet = self.sheet,

                map_row_height = sheet.rowHeights || {},

                list_row_height_index = [],
                list_row_height_offset = [],

                list_row_top = [],
                list_row_top_index = [],

                offset_height = 0,

                list_row_offset = [self._colHeadHeight],
                def_row_height = self._rowHeight,
                i;

            var d1 = Date.now();

            Object.keys(map_row_height).sort(function (a, b) {
                return a - b;
            }).forEach(function (key, i) {
                var row = +key;
                list_row_height_index.push(row);

                offset_height += map_row_height[key];
                list_row_height_offset.push(offset_height);

                if (i == 0) {
                    list_row_top.push(row * def_row_height);
                    list_row_top_index.push(row);
                }
                list_row_top.push(offset_height + (row - i) * def_row_height);
                list_row_top_index.push(row + 1);
            });

            var d2 = Date.now();
            log("updateRowsOffset: " + (d2 - d1) + "ms");

            self._rowHeightIndexs = list_row_height_index;
            self._rowHeightOffsets = list_row_height_offset;
            self._rowOffsets = list_row_offset;

            self._rowTops = list_row_top;
            self._rowTopIndexs = list_row_top_index;

            //预缓存数据
            //for (i = 1; i < 1000; i++) {
            //    list_row_offset[i] = (map_row_height[i - 1] || def_row_height) + list_row_offset[i - 1];
            //}

            return self;
        },

        getColLeft: function (col) {
            var self = this,
                value = self._colOffsets[col];

            if (value == undefined) {
                self._colOffsets[col] = value = calcCellOffset(self._colWidthIndexs, self._colWidthOffsets, self._colWidth, col) + self._rowNumWidth;
            }

            return value;
        },
        getRowTop: function (row) {
            var self = this,
                value = self._rowOffsets[row];

            if (value == undefined) {
                self._rowOffsets[row] = value = calcCellOffset(self._rowHeightIndexs, self._rowHeightOffsets, self._rowHeight, row) + self._colHeadHeight;
            }

            return value;
        },

        getCellLeft: function (col) {
            var self = this,
                left = self.getColLeft(col);

            if (col >= self._splitCol) left -= self.FIXED_WIDTH;

            return left;

            //return this.getColLeft(col);
        },

        getCellTop: function (row) {
            var self = this,
                top = self.getRowTop(row);

            if (row >= self._splitRow) top -= self.FIXED_HEIGHT;

            return top;

            //return this.getRowTop(row);
        },

        getColWidth: function (col) {
            return this.sheet.colWidths[col] || this._colWidth;
        },

        getColsWidth: function (firstCol, lastCol) {
            return firstCol != lastCol ? this.getColLeft(lastCol + 1) - this.getColLeft(firstCol) : this.getColWidth(firstCol);
        },

        getRowHeight: function (row) {
            return this.sheet.rowHeights[row] || this._rowHeight;
        },

        getRowsHeight: function (firstRow, lastRow) {
            return firstRow != lastRow ? this.getRowTop(lastRow + 1) - this.getRowTop(firstRow) : this.getRowHeight(firstRow);
        },

        findRow: function (top) {
            var self = this;

            return findCellIndex(self._rowTops, self._rowTopIndexs, self._rowHeight, top - self._colHeadHeight);
        },

        findCol: function (left) {
            var self = this;

            return findCellIndex(self._colLefts, self._colLeftIndexs, self._colWidth, left - self._rowNumWidth);
        },

        setColWidth: function (col, width) {
            var self = this,
                sheet = self.sheet,
                splitCol = self._splitCol;

            sheet.colWidths[col] = width;

            self.updateColsOffset();

            if (col < splitCol) elGrid3.style.width = elGrid1.style.width = self.getColLeft(splitCol) + "px";


        },

        setRowHeight: function (row, height) {

        },

        getTotalWidth: function () {
            return this.getColsWidth(0, this._maxCol + 1);
        },

        getTotalHeight: function () {
            return this.getRowsHeight(0, this._maxRow + 1);
        },

        getHScrollWidth: function () {
            var self = this;
            return self._HScrollbarWidth + self.getColsWidth(self._splitCol, self._maxCol + 1) - self.SCROLL_WIDTH;
        },

        getVScrollHeight: function () {
            var self = this;
            return self._VScrollbarHeight + self.getRowsHeight(self._splitRow, self._maxRow + 1) - self.SCROLL_HEIGHT;
        },

        updateHScrollWidth: function () {
            var self = this,
                elHScroll = self.elHScroll;

            if (elHScroll) getFirst(elHScroll).style.width = self.getHScrollWidth() + "px";

            return self;
        },

        updateVScrollHeight: function () {
            var self = this,
                elVScroll = self.elVScroll;

            if (elVScroll) {
                var total_height = self.getVScrollHeight();
                getFirst(elVScroll).style.height = total_height + "px";

                //修复IE8下垂直滚动条不显示的问题
                if (Q.ie < 8) elVScroll.style.overflowY = total_height > self._VScrollbarHeight ? "auto" : "scroll";
            }

            return self;
        },

        updateGridSettings: function () {
            var self = this,
                sheet = self.sheet,

                scrollRow = sheet.scrollRow,
                scrollCol = sheet.scrollCol,

                isFreeze = sheet.isFreeze;

            self._colWidth = sheet.colWidth || DEF_COL_WIDTH;
            self._rowHeight = sheet.rowHeight || DEF_ROW_HEIGHT;

            self._splitRow = isFreeze ? sheet.freezeRow || 0 : 0;
            self._splitCol = isFreeze ? sheet.freezeCol || 0 : 0;

            self._pageRowsNum = Math.ceil(self.GRID_HEIGHT / self._rowHeight) + 10 - self._splitRow;
            self._pageColsNum = Math.ceil(self.GRID_WIDTH / self._colWidth) + 5 - self._splitCol;

            self._maxRowsNum = self._pageRowsNum * 2;
            self._maxColsNum = self._pageColsNum * 2;

            self._drawFirstRow = Math.max(self._splitRow, scrollRow ? scrollRow - 5 : 0);
            self._drawFirstCol = Math.max(self._splitCol, scrollCol ? scrollCol - 2 : 0);
            self._drawLastRow = self._drawFirstRow + self._pageRowsNum;
            self._drawLastCol = self._drawFirstCol + self._pageColsNum;

            self._maxRow = Math.max(sheet.maxRow, self._drawLastRow);
            self._maxCol = Math.max(sheet.maxCol, self._drawLastCol);

            self._fixLine = sheet.displayGridlines ? -1 : 0;

            self._scrollRow = self._splitRow;
            self._scrollCol = self._splitCol;

            self._scrollLeft = 0;
            self._scrollTop = 0;

            //test
            //self._pageRowsNum = 60;
            //self._pageColsNum = 32;

            return self;
        },

        updateGridMerges: function () {
            var self = this,
                sheet = self.sheet,
                mapMerge = {};

            sheet.merges.forEach(function (data) {
                for (var i = data[0], lastRow = data[2]; i <= lastRow; i++) {
                    if (!mapMerge[i]) mapMerge[i] = {};

                    for (var j = data[1], lastCol = data[3]; j <= lastCol ; j++) {
                        mapMerge[i][j] = data;
                    }
                }
            });

            self.mapMerge = mapMerge;

            return self;
        },

        getMergeData: function (row, col) {
            //INFINITY_VALUE 表示整行或整列
            var mapMerge = this.mapMerge,
                mapRow = mapMerge[INFINITY_VALUE] || mapMerge[row];

            return mapRow ? mapRow[INFINITY_VALUE] || mapRow[col] : undefined;
        },

        getMergeState: function (firstRow, firstCol, lastRow, lastCol) {
            var self = this,
                data = self.getMergeData(firstRow, firstCol);

            if (!data) return RANGE_NO_MERGE;

            var isFirst = data[0] == firstRow && data[1] == firstCol;
            if (isFirst && data[2] == lastRow && data[3] == lastCol) return RANGE_IS_MERGE;

            return isFirst ? RANGE_IS_MERGE_FIRST : RANGE_IS_MERGE_OTHER;
        },

        isMerge: function (firstRow, firstCol, lastRow, lastCol) {
            return this.getMergeState(firstRow, firstCol, lastRow, lastCol) == RANGE_IS_MERGE;
        },

        //按区域划分合并数据([区域内,区域外])
        splitAllMerges: function (firstRow, firstCol, lastRow, lastCol) {
            var self = this,
                dataMerges = self.sheet.merges;

            if (firstRow == INFINITY_VALUE && firstCol == INFINITY_VALUE) return [dataMerges, []];

            var inRegionList = [], outRegionList = [], i = 0, len = dataMerges.length, data;

            if (firstRow == INFINITY_VALUE) {
                for (; i < len; i++) {
                    data = dataMerges[i];
                    if (firstCol <= data[1] && lastCol >= data[3]) inRegionList.push(data);
                    else outRegionList.push(data);
                }
            } else if (firstCol == INFINITY_VALUE) {
                for (; i < len; i++) {
                    data = dataMerges[i];
                    if (firstRow <= data[0] && lastRow >= data[2]) inRegionList.push(data);
                    else outRegionList.push(data);
                }
            } else {
                for (; i < len; i++) {
                    data = dataMerges[i];

                    if (firstRow <= data[0] && firstCol <= data[1] && lastRow >= data[2] && lastCol >= data[3]) inRegionList.push(data);
                    else outRegionList.push(data);
                }
            }

            return [inRegionList, outRegionList];
        },

        //计算选区
        computeRange: function (firstRow, firstCol, lastRow, lastCol) {
            return computeRange(this.sheet.merges, firstRow, firstCol, lastRow, lastCol);
        },

        _mergeCells2: function (firstRow, firstCol, lastRow, lastCol) {
            var self = this,
                list_merge_cell = [];

            self.showPanel();
        },

        getElMerges: function (data) {
            return self.mapElMerges[data.join("_")];
        },

        _mergeCells: function (data) {
            var self = this,

                _firstRow = data[0],
                _firstCol = data[1],
                _lastRow = data[2],
                _lastCol = data[3],

                key = data.join("_"),
                mapElMerges = self.mapElMerges,
                elMerges = mapElMerges[key];

            if (!elMerges) {
                mapElMerges[key] = elMerges = [];

                var node = createEle("div", "ss-merge"), i = 0, el;
                while (i < 4) {
                    el = node.cloneNode(true);
                    elMerges[i++] = el;
                    self["elMerges" + i].appendChild(el);
                }
            }

            var firstRow = data[0],
                firstCol = data[1],
                lastRow = data[2],
                lastCol = data[3];

            if (firstRow == INFINITY_VALUE) {
                firstRow = 0;
                lastRow = self._drawLastRow;
            }

            if (firstCol == INFINITY_VALUE) {
                firstCol = 0;
                lastCol = self._drawLastCol;
            }

            var left = self.getColLeft(firstCol),
                top = self.getRowTop(firstRow),
                width = self.getColsWidth(firstCol, lastCol),
                height = self.getRowsHeight(firstRow, lastRow),

                cssText = 'left:' + left + 'px;top:' + top + 'px;width:' + (width - 1) + 'px;height:' + (height - 1) + 'px;',
                cssFixTop = 'margin-top: -' + self.FIXED_HEIGHT + 'px;',
                cssFixLeft = 'margin-left: -' + self.FIXED_WIDTH + 'px;';

            elMerges[0].style.cssText = cssText;
            elMerges[1].style.cssText = cssText + cssFixLeft;
            elMerges[2].style.cssText = cssText + cssFixTop;
            elMerges[3].style.cssText = cssText + cssFixLeft + cssFixTop;

            return self;
        },

        mergeCells: function (firstRow, firstCol, lastRow, lastCol) {
            return this._mergeCells([firstRow, firstCol, lastRow, lastCol]);
        },

        initGridMerges: function () {
            var self = this,
                sheet = self.sheet;

            var elMerges1 = createEle("div", "ss-merges"),
                elMerges2 = elMerges1.cloneNode(true),
                elMerges3 = elMerges1.cloneNode(true),
                elMerges4 = elMerges1.cloneNode(true);

            self.elGrid1.appendChild(elMerges1);
            self.elGrid2.appendChild(elMerges2);
            self.elGrid3.appendChild(elMerges3);
            self.elGrid4.appendChild(elMerges4);

            self.elMerges1 = elMerges1;
            self.elMerges2 = elMerges2;
            self.elMerges3 = elMerges3;
            self.elMerges4 = elMerges4;

            self.updateGridMerges();

            self.mapElMerges = {};

            sheet.merges.forEach(function (data) {
                self._mergeCells(data);
            });

            return self;
        },

        updateGridSize: function () {
            var self = this,

                elGrid1 = self.elGrid1,
                elGrid2 = self.elGrid2,
                elGrid3 = self.elGrid3,
                elGrid4 = self.elGrid4,

                elHScroll = self.elHScroll,
                elVScroll = self.elVScroll,

                splitRow = self._splitRow,
                splitCol = self._splitCol,

                GRID_WIDTH = self.GRID_WIDTH,
                GRID_HEIGHT = self.GRID_HEIGHT,

                FIXED_WIDTH = self.getColLeft(splitCol),
                FIXED_HEIGHT = self.getRowTop(splitRow);

            self.FIXED_WIDTH = FIXED_WIDTH;
            self.FIXED_HEIGHT = FIXED_HEIGHT;

            self.SCROLL_WIDTH = GRID_WIDTH - FIXED_WIDTH;
            self.SCROLL_HEIGHT = GRID_HEIGHT - FIXED_HEIGHT;

            //elGrid4.style.left = elGrid2.style.left = elGrid3.style.width = elGrid1.style.width = FIXED_WIDTH + "px";
            //elGrid4.style.top = elGrid3.style.top = elGrid2.style.height = elGrid1.style.height = FIXED_HEIGHT + "px";

            elGrid3.style.width = elGrid1.style.width = FIXED_WIDTH + (splitCol ? 0 : 2) + "px";
            elGrid2.style.height = elGrid1.style.height = FIXED_HEIGHT + (splitRow ? 0 : 2) + "px";

            elGrid4.style.left = elGrid2.style.left = FIXED_WIDTH + "px";
            elGrid4.style.top = elGrid3.style.top = FIXED_HEIGHT + "px";

            elGrid4.style.width = elGrid2.style.width = self.SCROLL_WIDTH + "px";
            elGrid4.style.height = elGrid3.style.height = self.SCROLL_HEIGHT + "px";

            //elGrid3.style.width = elGrid1.style.width = FIXED_WIDTH + "px";
            //elGrid4.style.width = elGrid2.style.width = GRID_WIDTH + "px";

            //elGrid2.style.height = elGrid1.style.height = FIXED_HEIGHT + "px";
            //elGrid4.style.height = elGrid3.style.height = GRID_HEIGHT + "px";

            self.updateHScrollWidth();

            if (elVScroll) {
                elVScroll.style.height = GRID_HEIGHT + "px";

                self._VScrollbarHeight = GRID_HEIGHT;
                self.updateVScrollHeight();
            }

            return self;
        },

        getElRowNum: function (row) {
            return this.elRowNums[row];
        },
        getElColHead: function (col) {
            return this.elColHeads[col];
        },
        getElRowCells: function (row) {
            return this.elCells[row];
        },
        getElCell: function (row, col) {
            var elRowCells = this.getElRowCells(row);
            return elRowCells ? elRowCells[col] : undefined;
        },

        getElSSRowLeft: function (row) {
            var elRowNum = this.getElRowNum(row);
            if (elRowNum) return elRowNum.parentNode;
        },
        getElSSRowRight: function (row) {
            var elRowCells = this.getElRowCells(row);
            if (elRowCells) {
                var elCell = elRowCells[elRowCells.length - 1];
                if (elCell) return elCell.parentNode;
            }
        },

        //预处理行高、列宽、偏移
        processCssOffset: function (firstRow, lastRow, firstCol, lastCol, splitRow, splitCol) {
            var self = this,

                fixLine = self._fixLine,

                list_css_left = [],
                list_css_top = [],
                list_css_width = [],
                list_css_height = [],
                list_css_inner_width = [],
                list_css_inner_height = [],

                i,
                j;

            //预处理列宽、偏移
            var process_col_offset = function (firstCol, lastCol) {
                for (j = firstCol; j < lastCol; j++) {
                    var width = self.getColWidth(j);
                    list_css_left[j] = "left:" + self.getCellLeft(j) + "px;";
                    list_css_width[j] = "width:" + width + "px;";
                    list_css_inner_width[j] = "width:" + (width + fixLine) + "px;";
                }
            };

            //预处理行高、偏移
            var process_row_offset = function (firstRow, lastRow) {
                for (i = firstRow; i < lastRow; i++) {
                    var height = self.getRowHeight(i);
                    list_css_top[i] = "top:" + self.getCellTop(i) + "px;";
                    list_css_height[i] = "height:" + height + "px;";
                    list_css_inner_height[i] = "height:" + (height + fixLine) + "px;";
                }
            };

            if (splitCol) process_col_offset(0, splitCol);
            process_col_offset(firstCol, lastCol);

            if (splitRow) process_row_offset(0, splitRow);
            process_row_offset(firstRow, lastRow);

            return {
                cssLeft: list_css_left,
                cssWidth: list_css_width,
                cssInnerWidth: list_css_inner_width,

                cssTop: list_css_top,
                cssHeight: list_css_height,
                cssInnerHeight: list_css_inner_height,

                cssRowNumWidth: "width:" + self._rowNumWidth + "px;",
                cssRowNumInnerWidth: "width:" + (self._rowNumWidth + fixLine) + "px;"
            };
        },

        drawRows: function (firstRow, lastRow) {
            var self = this,
                sheet = self.sheet,

                firstCol = self._drawFirstCol,
                lastCol = self._drawLastCol + 1,

                splitRow = self._splitRow,
                splitCol = self._splitCol,

                i,
                j;

            //预处理行高、列宽、偏移等数据
            var pco = self.processCssOffset(firstRow, lastRow, firstCol, lastCol, 0, splitCol),
                list_css_left = pco.cssLeft,
                list_css_top = pco.cssTop,
                list_css_width = pco.cssWidth,
                list_css_height = pco.cssHeight,
                list_css_inner_width = pco.cssInnerWidth,
                list_css_inner_height = pco.cssInnerHeight,

                css_rowNum_width = pco.cssRowNumWidth,
                css_rowNum_inner_width = pco.cssRowNumInnerWidth;

            //画单元格（包括行号）
            var draw_cells = function (firstRow, lastRow, firstCol, lastCol, hasRowNum) {
                var tmp = [];
                for (i = firstRow; i < lastRow; i++) {
                    tmp.push('<div class="ss-row" style="' + list_css_top[i] + list_css_height[i] + '">');
                    if (hasRowNum) {
                        tmp.push('<div class="ss-rowNum" style="left:0px;' + css_rowNum_width + '">');
                        tmp.push('<div class="ss-val" style="' + css_rowNum_inner_width + list_css_inner_height[i] + '">');
                        tmp.push(i + 1);
                        tmp.push('</div>');
                        tmp.push('</div>');
                    }
                    for (j = firstCol; j < lastCol; j++) {
                        tmp.push('<div class="ss-cell" style="' + list_css_left[j] + list_css_width[j] + '">');
                        tmp.push('<div class="ss-val" style="' + list_css_inner_width[j] + list_css_inner_height[i] + '"></div>');
                        tmp.push('</div>');
                    }
                    tmp.push('</div>');
                }

                return tmp.join('');
            };

            var boxLeft = document.createElement("div"),
                boxRight = document.createElement("div"),

                elFragLeft = document.createDocumentFragment(),
                elFragRight = document.createDocumentFragment(),

                node;

            boxLeft.innerHTML = draw_cells(firstRow, lastRow, 0, splitCol, true);
            boxRight.innerHTML = draw_cells(firstRow, lastRow, firstCol, lastCol, false);

            while ((node = boxLeft.firstChild)) {
                elFragLeft.appendChild(node);
            }

            while ((node = boxRight.firstChild)) {
                elFragRight.appendChild(node);
            }

            var elGridLeft = getFirst(self.elGridCells3),
                elGridRight = getFirst(self.elGridCells4);

            if (firstRow > self._drawLastRow) {
                elGridLeft.appendChild(elFragLeft);
                elGridRight.appendChild(elFragRight);
            } else {
                elGridLeft.insertBefore(elFragLeft, elGridLeft.firstChild);
                elGridRight.insertBefore(elFragRight, elGridRight.firstChild);
            }

            return self;
        },

        drawCols: function (firstCol, lastCol) {
            var self = this,
                sheet = self.sheet,

                firstRow = self._drawFirstRow,
                lastRow = self._drawLastRow + 1,

                drawFirstCol = self._drawFirstCol,
                drawLastCol = self._drawLastCol,

                splitRow = self._splitRow,
                splitCol = self._splitCol,

                elColHeads = self.elColHeads,
                elCells = self.elCells,

                i,
                j;

            //预处理行高、列宽、偏移等数据
            var pco = self.processCssOffset(firstRow, lastRow, firstCol, lastCol, splitRow, 0),
                list_css_left = pco.cssLeft,
                list_css_top = pco.cssTop,
                list_css_width = pco.cssWidth,
                list_css_height = pco.cssHeight,
                list_css_inner_width = pco.cssInnerWidth,
                list_css_inner_height = pco.cssInnerHeight;

            var append_node = function (els, elFrag) {
                var nextNode;

                if (firstCol > drawLastCol) {
                    nextNode = els[drawLastCol];
                    if (nextNode) nextNode.parentNode.appendChild(elFrag);
                } else {
                    nextNode = els[drawFirstCol];
                    if (nextNode) nextNode.parentNode.insertBefore(elFrag, nextNode);
                }
            };

            var append_cols = function (row) {
                var elFragCells = document.createDocumentFragment(), nextNode;

                for (j = firstCol; j < lastCol; j++) {
                    var el = createEle("div", "ss-cell", '<div class="ss-val" style="' + list_css_inner_width[j] + list_css_inner_height[i] + '"></div>');
                    el.style.cssText = list_css_left[j] + list_css_width[j];
                    elFragCells.appendChild(el);
                }

                append_node(elCells[row], elFragCells);
            };

            //添加表头
            var elFragColHead = document.createDocumentFragment();

            for (j = firstCol; j < lastCol; j++) {
                var el = createEle("div", "ss-colHead", '<div class="ss-val" style="' + list_css_inner_width[j] + '">' + getColName(j) + '</div>');
                el.style.cssText = list_css_left[j] + list_css_width[j];
                elFragColHead.appendChild(el);
            }

            append_node(elColHeads, elFragColHead);

            for (i = 0; i < splitRow; i++) {
                append_cols(i);
            }

            for (i = firstRow; i < lastRow; i++) {
                append_cols(i);
            }

            return self;
        },

        updateElCache: function () {
            var self = this,

                elSheet = self.elSheet,

                split_row = self._splitRow,
                split_col = self._splitCol,

                draw_first_row = self._drawFirstRow,
                draw_first_col = self._drawFirstCol,

                elRowNums = [],
                elColHeads = [],
                elCells = [],

                nodes1 = self.elGridCells1.childNodes,
                nodes2 = self.elGridCells2.childNodes,
                nodes3 = self.elGridCells3.childNodes,
                nodes4 = self.elGridCells4.childNodes,

                nodes_head1 = nodes1[0].childNodes,
                nodes_head2 = nodes2[0].childNodes,

                len_rows1 = nodes1.length,
                len_rows2 = nodes3.length,

                len_cols1 = nodes_head1.length,
                len_cols2 = nodes_head2.length,

                i;

            for (i = 1; i < len_cols1; i++) {
                elColHeads[i - 1] = nodes_head1[i];
            }

            for (i = 0; i < len_cols2; i++) {
                elColHeads[draw_first_col + i] = nodes_head2[i];
            }

            var walk_row = function (row1, row2, i) {
                var nodes1 = row1.childNodes,
                    nodes2 = row2.childNodes,
                    len1 = nodes1.length,
                    len2 = nodes2.length,
                    list_cell = [],
                    j;

                for (j = 0; j < len1; j++) {
                    if (j == 0) elRowNums[i] = nodes1[j];
                    else list_cell[j - 1] = nodes1[j];
                }

                for (j = 0; j < len2; j++) {
                    list_cell[draw_first_col + j] = nodes2[j];
                }

                elCells[i] = list_cell;
            };

            for (i = 1; i < len_rows1; i++) {
                walk_row(nodes1[i], nodes2[i], i - 1);
            }

            for (i = 0; i < len_rows2; i++) {
                walk_row(nodes3[i], nodes4[i], draw_first_row + i);
            }

            self.elColStart = nodes_head1[0];
            self.elColHeads = elColHeads;
            self.elRowNums = elRowNums;
            self.elCells = elCells;

            return self;
        },

        setRowNum: function (row) {
            var el = this.getElRowNum(row), node;
            if (el) {
                node = el.firstChild;
                if (node) node.innerHTML = row + 1;
            }

            return this;
        },

        //更新行号
        updateRowNum: function (firstRow, lastRow) {
            var self = this,
                splitRow = self._splitRow,
                drawFirstRow = self._drawFirstRow,
                drawLastRow = self._drawLastRow,

                elRowNums = self.elRowNums,
                elCells = self.elCells,
                i;

            if (firstRow < splitRow) {
                for (i = firstRow; i < splitRow; i++) {
                    self.setRowNum(i);
                }
            }

            if (firstRow < drawFirstRow) firstRow = drawFirstRow;
            if (lastRow == -1 || lastRow > drawLastRow) lastRow = drawLastRow;

            for (i = firstRow; i <= lastRow; i++) {
                self.setRowNum(i);
            }

            return self;
        },

        //更新列名
        updateColHead: function (firstCol, lastCol) {

        },

        fillCells: function (firstRow, firstCol, lastRow, lastCol) {
            var self = this,
                sheet = self.sheet;

            return self;
        },

        scrollToRow: function (relRow, isUpdateScrollbar) {
            if (relRow < 0) relRow = 0;

            var self = this,

                scrollRow = self._scrollRow,

                drawFirstRow = self._drawFirstRow,
                drawLastRow = self._drawLastRow,
                splitRow = self._splitRow,

                row = relRow + splitRow,

                maxRow = self._maxRow,
                pageRowsNum = self._pageRowsNum,

                updated = false;

            var top = self.getRowTop(row) - self.getRowTop(splitRow),

                firstRow,
                lastRow;

            var doScrollRow = function (redraw, isup) {
                if (redraw) {
                    //必须在drawCells之前更新
                    self._drawFirstRow = firstRow;
                    self._drawLastRow = lastRow - 1;

                    self.drawCells();
                } else {
                    self.drawRows(firstRow, lastRow);

                    //必须在drawRows之后更新
                    if (isup) self._drawFirstRow = firstRow;
                    else self._drawLastRow = lastRow - 1;
                }
            };

            //向上滚动
            if (relRow < scrollRow) {
                if (row < drawFirstRow) {
                    firstRow = Math.max(splitRow, row);
                    lastRow = Math.min(firstRow + pageRowsNum, drawFirstRow);

                    doScrollRow(lastRow < drawFirstRow, true);

                    updated = true;
                }
            } else if (relRow > scrollRow) {
                //向下滚动

                var isOutOfHeight = self.getRowsHeight(row, maxRow) < self.SCROLL_HEIGHT;

                if (row > drawLastRow || isOutOfHeight) {
                    if (row > drawLastRow) {
                        firstRow = Math.max(row - pageRowsNum, drawLastRow + 1);
                        lastRow = row + 1;

                    } else {
                        firstRow = drawLastRow + 1;
                        lastRow = firstRow + relRow - scrollRow;
                    }

                    doScrollRow(firstRow > drawLastRow + 1 && !isOutOfHeight, false);

                    //更新垂直滚动条高度
                    if (lastRow - 1 > maxRow) {
                        self._maxRow = lastRow - 1;
                        self.updateVScrollHeight();
                    }

                    updated = true;
                }
            }

            if (updated) self.updateElCache();

            self.elGrid4.scrollTop = self.elGrid3.scrollTop = self._scrollTop = top;

            //self.updateElLineRowsActive(self.sheet.firstRow);

            if (isUpdateScrollbar) {
                self.skipScrollEvent = true;
                self.elVScroll.scrollTop = top;
            }

            self.sheet.scrollRow = self._scrollRow = relRow;

            return self;
        },
        scrollToCol: function (relCol, isUpdateScrollbar) {
            if (relCol < 0) relCol = 0;

            var self = this,

                scrollCol = self._scrollCol,

                drawFirstCol = self._drawFirstCol,
                drawLastCol = self._drawLastCol,
                splitCol = self._splitCol,

                col = relCol + splitCol,

                maxCol = self._maxCol,
                pageColsNum = self._pageColsNum,

                updated = false;

            var left = self.getColLeft(col) - self.getColLeft(splitCol),

                firstCol,
                lastCol;

            var doScrollCol = function (redraw, isleft) {
                if (redraw) {
                    //必须在drawCells之前更新
                    self._drawFirstCol = firstCol;
                    self._drawLastCol = lastCol - 1;

                    self.drawCells();
                } else {
                    self.drawCols(firstCol, lastCol);

                    //必须在drawCols之后更新
                    if (isleft) self._drawFirstCol = firstCol;
                    else self._drawLastCol = lastCol - 1;
                }
            };

            //向左滚动
            if (relCol < scrollCol) {
                if (col < drawFirstCol) {
                    firstCol = Math.max(splitCol, col);
                    lastCol = Math.min(firstCol + pageColsNum, drawFirstCol);

                    doScrollCol(lastCol < drawFirstCol, true);

                    updated = true;
                }
            } else if (relCol > scrollCol) {
                //向右滚动

                var isOutOfWidth = self.getColsWidth(col, maxCol) < self.SCROLL_WIDTH;

                if (col > drawLastCol || isOutOfWidth) {
                    if (col > drawLastCol) {
                        firstCol = Math.max(col - pageColsNum, drawLastCol + 1);
                        lastCol = col + 1;

                    } else {
                        firstCol = drawLastCol + 1;
                        lastCol = firstCol + relCol - scrollCol;
                    }

                    doScrollCol(firstCol > drawLastCol + 1 && !isOutOfWidth, false);

                    //更新水平滚动条宽度
                    if (lastCol - 1 > maxCol) {
                        self._maxCol = lastCol - 1;
                        self.updateHScrollWidth();
                    }

                    updated = true;
                }
            }

            if (updated) self.updateElCache();

            self.elGrid4.scrollLeft = self.elGrid2.scrollLeft = self._scrollLeft = left;

            //self.updateElLineColsActive(self.sheet.firstCol);

            if (isUpdateScrollbar) {
                self.skipScrollEvent = true;
                self.elHScroll.scrollLeft = left;
            }

            self.sheet.scrollCol = self._scrollCol = relCol;

            return self;
        },

        scrollToRowAsync: function (row, isUpdateScrollbar, callback) {
            var self = this;

            async(function () {
                self.scrollToRow(row, isUpdateScrollbar);
                fire(callback, self);
            }, 0);

            return self;
        },
        scrollToColAsync: function (col, isUpdateScrollbar, callback) {
            var self = this;

            async(function () {
                self.scrollToCol(col, isUpdateScrollbar);
                fire(callback, self);
            }, 0);

            return self;
        },

        activeColHeads: function (firstCol, lastCol) {
            var self = this;

            $(self.elColHeads.slice(firstCol, lastCol + 1)).addClass("ss-on");

            return self;
        },

        activeRowNums: function (firstRow, lastRow) {
            var self = this;

            $(self.elRowNums.slice(firstRow, lastRow + 1)).addClass("ss-on");

            return self;
        },

        //updateElLineColsActive: function (firstCol) {
        //    var self = this,
        //        elLineColsActive = self.elLineColsActive;

        //    if (elLineColsActive) elLineColsActive.style.left = (self.getColLeft(Math.max(firstCol, 0)) - self._scrollLeft) + "px";

        //    return self;
        //},

        //updateElLineRowsActive: function (firstRow, firstCol) {
        //    var self = this,
        //        elLineRowsActive = self.elLineRowsActive;

        //    if (elLineRowsActive) elLineRowsActive.style.top = (self.getRowTop(Math.max(firstRow, 0)) - self._scrollTop) + "px";

        //    return self;
        //},

        showSplitHLine: function (row) {
            var self = this,
                elSplitHLine = self.elSplitHLine;

            if (!row) return self;

            if (!elSplitHLine) {
                self.elSplitHLine = elSplitHLine = createEle("div", "ss-split-hline");
                self.elSheet.appendChild(elSplitHLine);
            }

            elSplitHLine.style.cssText = "left:0px;top:" + (self.getRowTop(row) - 1) + "px;width:" + self.GRID_WIDTH + "px;";

            return self;

        },
        showSplitVLine: function (col) {
            var self = this,
                elSplitVLine = self.elSplitVLine;

            if (!col) return self;

            if (!elSplitVLine) {
                self.elSplitVLine = elSplitVLine = createEle("div", "ss-split-vline");
                self.elSheet.appendChild(elSplitVLine);
            }

            elSplitVLine.style.cssText = "left:" + (self.getColLeft(col) - 1) + "px;top:0px;height:" + self.GRID_HEIGHT + "px;";

            return self;
        },

        showSplitLine: function () {
            var self = this;

            return self.showSplitHLine(self._splitRow).showSplitVLine(self._splitCol);
        },

        showDragHLine: function (row) {
            var self = this,
                elDragHLine = self.elDragHLine;

            if (!row) return self;

            if (!elDragHLine) {
                self.elDragHLine = elDragHLine = createEle("div", "ss-drag-hline");
                self.elSheet.appendChild(elDragHLine);
            }

            elDragHLine.style.cssText = "left:0px;top:" + (self.getRowTop(row) - 1) + "px;width:" + self.GRID_WIDTH + "px;";

            return self;

        },
        showDragVLine: function (col) {
            var self = this,
                elDragVLine = self.elDragVLine;

            if (!col) return self;

            if (!elDragVLine) {
                self.elDragVLine = elDragVLine = createEle("div", "ss-drag-vline");
                self.elSheet.appendChild(elDragVLine);
            }

            elDragVLine.style.cssText = "left:" + (self.getColLeft(col) - 1) + "px;top:0px;height:" + self.GRID_HEIGHT + "px;";

            return self;
        },

        hideDragHLine: function () {
            var self = this,
                elDragHLine = self.elDragHLine;

            if (elDragHLine) elDragHLine.style.display = "none";

            return self;
        },
        hideDragVLine: function () {
            var self = this,
                elDragVLine = self.elDragVLine;

            if (elDragVLine) elDragVLine.style.display = "none";

            return self;
        },

        showPanel: function (firstRow, firstCol, lastRow, lastCol) {
            var self = this,

                sheet = self.sheet,

                splitRow = self._splitRow,
                splitCol = self._splitCol,

                elSheet = self.elSheet,
                elPanel1 = self.elPanel1,
                elPanel2 = self.elPanel2,
                elPanel3 = self.elPanel3,
                elPanel4 = self.elPanel4,

                elLineColsActive = self.elLineColsActive,
                elLineRowsActive = self.elLineRowsActive;

            if (lastRow == undefined) lastRow = firstRow;
            if (lastCol == undefined) lastCol = firstCol;

            if (!elPanel1) {
                self.elPanel1 = elPanel1 = createEle("div", "ss-panel", '<div class="ss-p-area"></div><div class="ss-p-resize"></div>');
                $(getFirst(elPanel1)).css("opacity", 0.14);

                self.elPanel2 = elPanel2 = elPanel1.cloneNode(true);
                self.elPanel3 = elPanel3 = elPanel1.cloneNode(true);
                self.elPanel4 = elPanel4 = elPanel1.cloneNode(true);

                self.elGrid1.appendChild(elPanel1);
                self.elGrid2.appendChild(elPanel2);
                self.elGrid3.appendChild(elPanel3);
                self.elGrid4.appendChild(elPanel4);

                //self.elLineColsActive = elLineColsActive = createEle("div", "ss-line-cols-active");
                //self.elLineRowsActive = elLineRowsActive = createEle("div", "ss-line-rows-active");

                //elSheet.appendChild(elLineColsActive);
                //elSheet.appendChild(elLineRowsActive);
            }

            if (firstRow == INFINITY_VALUE) {
                firstRow = 0;
                lastRow = self._drawLastRow;
            }

            if (firstCol == INFINITY_VALUE) {
                firstCol = 0;
                lastCol = self._drawLastCol;
            }

            $(".ss-on", elSheet).removeClass("ss-on");
            self.activeColHeads(firstCol, lastCol).activeRowNums(firstRow, lastRow);

            var left = self.getColLeft(firstCol),
                top = self.getRowTop(firstRow),
                width = self.getColsWidth(firstCol, lastCol),
                height = self.getRowsHeight(firstRow, lastRow),

                cssText = 'left:' + (left - 1) + 'px;top:' + (top - 1) + 'px;width:' + (width - 2) + 'px;height:' + (height - 2) + 'px;',
                cssFixTop = 'margin-top: -' + self.FIXED_HEIGHT + 'px;',
                cssFixLeft = 'margin-left: -' + self.FIXED_WIDTH + 'px;';

            //elLineColsActive.style.cssText = 'width:' + width + 'px;left:' + (left - self._scrollLeft) + 'px;top:' + (self.getRowTop(0) - 1) + 'px;';
            //elLineRowsActive.style.cssText = 'height:' + height + 'px;left:' + (self.getColLeft(0) - 1) + 'px;top:' + (top - self._scrollTop) + 'px;';

            elPanel1.style.cssText = cssText;
            elPanel2.style.cssText = cssText + cssFixLeft;
            elPanel3.style.cssText = cssText + cssFixTop;
            elPanel4.style.cssText = cssText + cssFixLeft + cssFixTop;

            return self;
        },

        selectCells: function (firstRow, firstCol, lastRow, lastCol) {
            var self = this,
                sheet = self.sheet,
                tmp;

            if (lastRow == undefined) lastRow = firstRow;
            if (lastCol == undefined) lastCol = firstCol;

            if (firstRow > lastRow) {
                tmp = firstRow;
                firstRow = lastRow;
                lastRow = tmp;
            }

            if (firstCol > lastCol) {
                tmp = firstCol;
                firstCol = lastCol;
                lastCol = tmp;
            }

            var isSelectOne = firstRow == lastRow && firstCol == lastCol;
            if (isSelectOne) {
                var merge = self.getMergeData(firstRow, firstCol);
                if (merge) {
                    if (firstRow > merge[0]) firstRow = merge[0];
                    if (firstCol > merge[1]) firstCol = merge[1];
                    if (lastRow < merge[2]) lastRow = merge[2];
                    if (lastCol < merge[3]) lastCol = merge[3];
                }
            } else {
                var range = self.computeRange(firstRow, firstCol, lastRow, lastCol);
                firstRow = range[0];
                firstCol = range[1];
                lastRow = range[2];
                lastCol = range[3];

                isSelectOne = self.isMerge(firstRow, firstCol, lastRow, lastCol);
            }

            self._isSelectOne = isSelectOne;

            self.showPanel(firstRow, firstCol, lastRow, lastCol);

            sheet.firstRow = firstRow;
            sheet.firstCol = firstCol;
            sheet.lastRow = lastRow;
            sheet.lastCol = lastCol;

            return self;
        },

        selectCellsAuto: function () {
            var self = this,
                sheet = self.sheet;

            return self.selectCells(sheet.firstRow, sheet.firstCol, sheet.lastRow, sheet.lastCol);
        },

        drawCells: function () {
            var self = this,
                sheet = self.sheet,

                tmp1 = [],
                tmp2 = [],
                tmp3 = [],
                tmp4 = [],

                fix_line = sheet.displayGridlines ? -1 : 0,

                draw_first_row = self._drawFirstRow,
                draw_first_col = self._drawFirstCol,
                draw_last_row = self._drawLastRow,
                draw_last_col = self._drawLastCol,

                split_row = self._splitRow,
                split_col = self._splitCol,

                i,
                j;

            var d1 = Date.now();

            //预处理行高、列宽、偏移等数据
            var pco = self.processCssOffset(draw_first_row, draw_last_row + 1, draw_first_col, draw_last_col + 1, split_row, split_col),
                list_css_left = pco.cssLeft,
                list_css_top = pco.cssTop,
                list_css_width = pco.cssWidth,
                list_css_height = pco.cssHeight,
                list_css_inner_width = pco.cssInnerWidth,
                list_css_inner_height = pco.cssInnerHeight,

                css_rowNum_width = pco.cssRowNumWidth,
                css_rowNum_inner_width = pco.cssRowNumInnerWidth;

            //画表头
            var draw_head = function (tmp, firstCol, lastCol, hasColStart) {
                tmp.push('<div class="ss-header">');
                if (hasColStart) tmp.push('<div class="ss-colStart" style="' + css_rowNum_width + '"><div class="ss-val" style="' + css_rowNum_inner_width + '"></div></div>');
                for (j = firstCol; j < lastCol; j++) {
                    //tmp.push('<div class="ss-colHead" style="' + list_css_left[j] + list_css_inner_width[j] + '">' + getColName(j) + '</div>');
                    tmp.push('<div class="ss-colHead" style="' + list_css_left[j] + list_css_width[j] + '">');
                    tmp.push('<div class="ss-val" style="' + list_css_inner_width[j] + '">');
                    tmp.push(getColName(j));
                    tmp.push('</div>');
                    tmp.push('</div>');
                }
                tmp.push('</div>');
            };

            //画单元格（包括行号）
            var draw_cells = function (tmp, firstRow, lastRow, firstCol, lastCol, hasRowNum) {
                for (i = firstRow; i < lastRow; i++) {
                    tmp.push('<div class="ss-row" style="' + list_css_top[i] + list_css_height[i] + '">');
                    if (hasRowNum) {
                        tmp.push('<div class="ss-rowNum" style="left:0px;' + css_rowNum_width + '">');
                        tmp.push('<div class="ss-val" style="' + css_rowNum_inner_width + list_css_inner_height[i] + '">');
                        tmp.push(i + 1);
                        tmp.push('</div>');
                        tmp.push('</div>');
                    }
                    for (j = firstCol; j < lastCol; j++) {
                        tmp.push('<div class="ss-cell" style="' + list_css_left[j] + list_css_width[j] + '">');
                        tmp.push('<div class="ss-val" style="' + list_css_inner_width[j] + list_css_inner_height[i] + '"></div>');
                        tmp.push('</div>');
                    }
                    tmp.push('</div>');
                }
            };

            //----------------- Grid1 -----------------
            draw_head(tmp1, 0, split_col, true);
            draw_cells(tmp1, 0, split_row, 0, split_col, true);

            //----------------- Grid2 -----------------
            draw_head(tmp2, draw_first_col, draw_last_col + 1, false);
            draw_cells(tmp2, 0, split_row, draw_first_col, draw_last_col + 1, false);

            //----------------- Grid3 -----------------
            draw_cells(tmp3, draw_first_row, draw_last_row + 1, 0, split_col, true);

            //----------------- Grid4 -----------------
            draw_cells(tmp4, draw_first_row, draw_last_row + 1, draw_first_col, draw_last_col + 1, false);

            var d2 = Date.now();

            self.elGridCells1.innerHTML = tmp1.join('');
            self.elGridCells2.innerHTML = tmp2.join('');
            self.elGridCells3.innerHTML = tmp3.join('');
            self.elGridCells4.innerHTML = tmp4.join('');

            var d3 = Date.now();
            log("drawCells: build html " + (d2 - d1) + "ms , render " + (d3 - d2) + "ms");

            return self.fillCells(draw_first_row, draw_first_col, draw_last_row, draw_last_col);
        },

        drawSheet: function () {
            var self = this,
                sheet = self.sheet;

            //test
            extend(sheet, { isFreeze: true, freezeRow: 3, freezeCol: 3, scrollRow: 0, scrollCol: 0, firstRow: 1, firstCol: 4, lastRow: 6, lastCol: 5, colWidths: { 1: 120, 7: 133, 8: 145, 9: 120, 12: 98, 20: 64, 38: 155 }, rowHeights: { 1: 67, 3: 45 }, merges: [[1, 1, 3, 2], [4, 5, 8, 6]] }, true);

            //sheet.isFreeze = false;

            self.updateGridSettings().updateColsOffset().updateRowsOffset().updateGridSize();

            async(function () {
                var d1 = Date.now();
                self.drawCells().updateElCache();
                var d2 = Date.now();

                document.title = "draw:" + (d2 - d1) + "ms";

                //避免浏览器缓存
                var elVScroll = self.elVScroll,
                    elHScroll = self.elHScroll;

                if (elVScroll) elVScroll.scrollTop = 1;
                if (elHScroll) elHScroll.scrollLeft = 1;

                self.loaded = true;

                //滚动到指定行和列
                if (sheet.scrollRow) self.scrollToRow(sheet.scrollRow, true);
                if (sheet.scrollCol) self.scrollToCol(sheet.scrollCol, true);

                self.showSplitLine().initGridMerges().selectCellsAuto();
            });

            return self;
        },

        fillSheet: function (sheet) {

        }
    });

    UI.register("initGrid");

})(window);