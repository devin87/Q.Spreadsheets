/*
* Q.UI.Box.js (包括遮罩层、拖动、弹出框)
* https://github.com/devin87/Q.UI.js
* author:devin87@qq.com
* update:2016/01/12 16:02
*/
(function (undefined) {
    "use strict";

    var window = Q.G,
        document = window.document,

        isObject = Q.isObject,
        isFunc = Q.isFunc,
        isUNum = Q.isUNum,
        isUInt = Q.isUInt,

        def = Q.def,
        fire = Q.fire,
        async = Q.async,
        extend = Q.extend,

        makeArray = Q.makeArray,

        getStyle = Q.getStyle,
        setStyle = Q.setStyle,

        getOffset = Q.offset,
        //setOffset = getOffset,

        getFirst = Q.getFirst,
        getNext = Q.getNext,
        //getLast = Q.getLast,

        setWidth = Q.width,
        setHeight = Q.height,

        createEle = Q.createEle,
        removeEle = Q.removeEle,

        cssShow = Q.show,
        cssHide = Q.hide,
        isHidden = Q.isHidden,
        cssToggle = Q.toggle,

        addClass = Q.addClass,

        setCssIfNot = Q.setCssIfNot,
        setCenter = Q.setCenter,

        setInputError = Q.setInputError,
        setInputDefault = Q.setInputDefault,
        clearSelection = Q.clearSelection,

        query = Q.query,

        factory = Q.factory,

        browser_ie = Q.ie,
        isIE6 = browser_ie < 7,

        view = Q.view,

        E = Q.event;

    //------------------------- Mask -------------------------

    var GLOBAL_MASK_SETTINGS = { color: "#999", opacity: 0.3 };

    //构造器:遮罩层
    function MaskBox(ops) {
        var box = createEle("div", "x-mask");
        Q.body.appendChild(box);

        if (isIE6) box.style.height = view.getScrollHeight() + "px";

        this.box = box;

        this.set(extend(ops || {}, GLOBAL_MASK_SETTINGS));
    }

    factory(MaskBox).extend({
        //设置透明度
        opacity: function (opacity) {
            return this.set({ opacity: opacity });
        },
        //设置遮罩层样式 eg:{color,zIndex,width,height}
        set: function (ops) {
            var box = this.box;

            Object.forEach(ops, function (key, value) {
                if (value !== undefined) setStyle(box, key != "color" ? key : "backgroundColor", value);
            });

            return this;
        },
        //显示遮罩层
        show: function () {
            if (this.removed) {
                Q.body.appendChild(this.box);
                this.removed = false;
            }

            cssShow(this.box);

            return this;
        },
        //隐藏遮罩层
        hide: function () {
            cssHide(this.box);

            return this;
        },
        //移除遮罩层
        remove: function () {
            removeEle(this.box);
            this.removed = true;

            return this;
        }
    });

    var GLOBAL_MASK_BOX;

    //获取通用遮罩层(默认配置)
    function getMaskBox() {
        if (GLOBAL_MASK_BOX) {
            GLOBAL_MASK_BOX.show();
            return GLOBAL_MASK_BOX;
        }

        GLOBAL_MASK_BOX = new MaskBox();
        return GLOBAL_MASK_BOX;
    }

    //修改遮罩层全局配置
    function maskSetup(ops, isUpdateUI) {
        extend(GLOBAL_MASK_SETTINGS, ops, true);

        if (isUpdateUI && GLOBAL_MASK_BOX) GLOBAL_MASK_BOX.reset(ops);
    }

    //------------------------- Drag -------------------------

    //获取所有顶层窗口(parent)
    function getTops(w) {
        var parents = [], cwin = w || window;
        while (cwin.top != cwin) {
            var pwin = cwin.parent;
            parents.push(pwin);
            cwin = pwin;
        }
        return parents;
    }

    /*//获取所有内嵌框架
    function _getFrames(w, ret) {
        var frames = w.frames, len = frames.length;
        if (len <= 0) return;
        for (var i = 0; i < len; i++) {
            ret.push(frames[i]);
            _getFrames(frames[i], ret);
        }
    }

    //获取所有内嵌框架
    function getFrames(w) {
        var frames = [];
        _getFrames(w || window, frames);
        return frames;
    }*/

    var DEF_INDEX = 1000, CURRENT_INDEX = DEF_INDEX, MASK_FOR_FRAME;

    //基础拖动对象
    function DragX(init) {
        var self = this;

        fire(init, self);

        self._apis = [];

        self.start();
    }

    factory(DragX).extend({
        //开始拖动
        start: function () {
            var self = this,
                ops = self.ops,

                win = ops.scope || window,
                document = win.document,

                ele = ops.ele,
                target = ops.target || ele,

                autoIndex = ops.autoIndex !== false,
                autoMask = ops.autoMask !== false,

                zIndex = ele.nodeType == 1 ? +getStyle(ele, "z-index") : 0,

                doDown = self.doDown,
                doMove = self.doMove,
                doUp = self.doUp,

                onCheck = ops.onCheck,

                onDown = ops.onDown,
                onMove = ops.onMove,
                onUp = ops.onUp,

                topWins = getTops(win),
                //frameWins,

                hasCapture = !!ele.setCapture,
                hasListen = !hasCapture && topWins.length > 0;  //是否需要监听其它窗口

            if (zIndex >= CURRENT_INDEX) CURRENT_INDEX = zIndex + 1;

            //初始化元素状态
            setCssIfNot(ele, "position", "absolute");
            setCssIfNot(target, "cursor", "move");

            //设置元素居中
            if (ops.center) {
                setCenter(ele);

                self._api_resize = E.add(win, "resize", function () {
                    setCenter(ele);
                });
            }

            //鼠标按下事件
            var mousedown = function (e) {
                if (isFunc(onCheck) && onCheck.call(self, e) === false) return;

                if (autoIndex) {
                    var _zIndex = +getStyle(ele, "z-index") || 0;
                    if (_zIndex < CURRENT_INDEX) {
                        CURRENT_INDEX++;
                        ele.style.zIndex = CURRENT_INDEX;
                    }
                }

                self._unbind();

                if (hasCapture) {
                    ele.setCapture();
                    self._bind(ele, "losecapture", mouseup);
                } else {
                    self._bind(win, "blur", mouseup);
                }

                self._bind(document, {
                    "mousemove": mousemove,
                    "mouseup": mouseup
                });

                if (hasListen) self._bind(topWins, "mouseup", mouseup);

                //frameWins = getFrames(win);
                //var hasMask = autoMask && !hasCapture && frameWins.length > 0;

                //创建一个遮罩层,在遮罩层上拖动
                //1. 避免某些浏览器在经过iframe时鼠标事件响应不正常
                //2. 避免拖动过程产生文本选区
                if (autoMask) {
                    if (!MASK_FOR_FRAME) MASK_FOR_FRAME = new MaskBox({ color: null, opacity: null, zIndex: 999999 });
                    if (isIE6) MASK_FOR_FRAME.set({ height: view.getScrollHeight() });
                }

                //清除文本选区
                async(clearSelection, 20);

                fire(doDown, self, e);
                fire(onDown, self, e);
            };

            //鼠标移动事件
            var mousemove = function (e) {
                if (self._pause) return;

                fire(doMove, self, e);
                fire(onMove, self, e);
            };

            //鼠标释放事件
            var mouseup = function (e) {
                self._unbind();

                if (hasCapture) ele.releaseCapture();

                if (autoMask && MASK_FOR_FRAME) {
                    MASK_FOR_FRAME.remove();
                    MASK_FOR_FRAME = null;
                }

                clearSelection();

                fire(doUp, self, e);
                fire(onUp, self, e);
            };

            self._up = mouseup;
            self._api = E.add(target, "mousedown", mousedown);

            return self;
        },

        //绑定事件
        _bind: function () {
            this._apis.push(E.add.apply(E, arguments));
        },
        //清理事件绑定
        _unbind: function () {
            var apis = this._apis;

            if (apis.length > 0) {
                apis.forEach(function (api) { api.off(); });
                this._apis = [];
            }
        },

        //暂停拖动
        pause: function (flag) {
            var self = this;

            self._pause = flag;
            self._up && self._up();

            return self;
        },
        //停止拖动
        stop: function () {
            var self = this,
                api = self._api;

            self._unbind();

            if (api) {
                api.off();
                self._api = null;
            }

            return self;
        },

        destroy: function () {
            var self = this,
                api_resize = self._api_resize;

            self.stop();
            if (api_resize) {
                api_resize.off();
                self._api_resize = null;
            }

            return self;
        }
    });

    var ELE_DRAG_SHADOW;

    //设置拖动
    function setDrag(ele, ops) {
        return new DragX(function () {
            if (ele.nodeType != 1) {
                ops = ele;
                ele = ops.ele;
            } else {
                ops = ops || {};
                ops.ele = ele;
            }

            var base = this,

                range = ops.range || { x: 0, y: 0 },
                hasShadow = ops.shadow !== false,

                w = ele.offsetWidth,
                h = ele.offsetHeight,

                target = ele,

                startX, startY, _isX, _isY;

            if (hasShadow) {
                if (!ELE_DRAG_SHADOW) {
                    ELE_DRAG_SHADOW = createEle("div", "x-drag-shadow");
                    Q.body.appendChild(ELE_DRAG_SHADOW);
                }

                target = ELE_DRAG_SHADOW;
            }

            //实现ops接口
            base.ops = ops;

            //实现doDown接口
            base.doDown = function (e) {
                startX = e.clientX - ele.offsetLeft;
                startY = e.clientY - ele.offsetTop;

                if (hasShadow) {
                    var offset = getOffset(ele);

                    Object.forEach({ left: offset.left, top: offset.top, width: ele.offsetWidth, height: ele.offsetHeight }, function (key, value) {
                        target.style[key] = value + "px";
                    });

                    //cssShow(target);
                }
            };

            //实现doMove接口
            base.doMove = function (e) {
                cssShow(target);

                if (_isX) {
                    var x = e.clientX - startX;
                    if (range) {
                        if (x < range.x) x = range.x;
                        else if (range.w && x + w > range.x + range.w) x = range.x + range.w - w;
                    }
                    target.style.left = x + "px";
                }
                if (_isY) {
                    var y = e.clientY - startY;
                    if (range) {
                        if (y < range.y) y = range.y;
                        else if (range.h && y + h > range.y + range.h) y = range.y + range.h - h;
                    }
                    target.style.top = y + "px";
                }
            };

            if (hasShadow) {
                base.doUp = function () {
                    if (isHidden(target)) return;
                    cssHide(target);

                    ele.style.left = target.style.left;
                    ele.style.top = target.style.top;
                };
            }

            //设置拖动方向
            base.setLock = function (isX, isY) {
                _isX = isX;
                _isY = isY;

                return base.pause(!isX && !isY);
            };

            //设置拖动范围
            base.setRange = function (x, y, w, h) {
                range = isObject(x) ? x : { x: x, y: y, w: w, h: h };

                return base;
            };

            base.setLock(ops.isX !== false, ops.isY !== false);
        });
    }

    $.fn.extend({
        drag: function (ops) {
            return this.each(function (i, el) {
                setDrag(el, ops);
            });
        }
    });

    //------------------------- Box -------------------------

    //接口,构造器:Box对象
    function Box(init) {
        this._es = [];

        fire(init, this);
    }

    factory(Box).extend({
        //在弹出框内查找对象
        find: function (pattern, context) {
            return typeof pattern == "string" ? query(pattern, context || this.box) : makeArray(pattern);
        },
        //获取弹出框内查找到的第一个对象
        get: function (pattern, context) {
            return this.find(pattern, context)[0];
        },
        //触发回调函数
        fire: function () {
            fire(this.callback, this, this.data);
            return this;
        },
        //获取事件回调函数
        getEventCallback: function (fn, data) {
            var self = this;

            if (fn == "hide") return function () { self.data = data; self.hide(); };
            if (fn == "remove") return function () { self.data = data; self.remove(); };

            return fn;
        },

        //绑定事件,同 Event.add,不过会缓存事件句柄,用于统一销毁
        bind: function () {
            this._es.push(E.add.apply(E, arguments));
            return this;
        },
        //将事件绑定到查找到的元素上,并非事件代理
        on: function (pattern, types, fn, data) {
            var self = this,
                list = self.find(pattern);

            if (isObject(types)) {
                data = fn;

                Object.forEach(types, function (type, fn) {
                    self.bind(list, type, self.getEventCallback(fn, data));
                });
            } else {
                self.bind(list, types, self.getEventCallback(fn, data));
            }

            return self;
        },
        //显示
        show: function () {
            var self = this;

            if (self.onShow) self.onShow();

            cssShow(self.box);
            if (self.mbox) self.mbox.show();

            return self;
        },
        //隐藏
        hide: function () {
            var self = this;

            cssHide(self.box);
            if (self.mbox) self.mbox.hide();
            if (self.onHide) self.onHide();

            return self.fire();
        },
        //自动切换显示或隐藏
        toggle: function () {
            cssToggle(this.box);
            return this;
        },
        //移除
        remove: function () {
            var self = this;
            if (!self.box) return;

            removeEle(self.box);

            //遮罩层
            if (self.mbox) self.mbox.remove();

            //拖动框
            if (self.dr) self.dr.destroy();

            //解除绑定的事件
            self._es.forEach(function (api) {
                api.off();
            });

            self.box = self.mbox = self.dr = null;

            if (self.onRemove) self.onRemove();

            return self.fire();
        }
    });

    Box.alias("remove", "destroy");

    //构造器:弹出层
    function WinBox(ops) {
        ops = ops || {};

        return new Box(function () {
            var self = this,

                width = ops.width,
                height = ops.height,
                maxHeight = ops.maxHeight,
                className = ops.className;

            self.callback = ops.callback;

            var html = '' +
                '<div class="x-head">' +
                    '<h2 class="x-title">' + (ops.title || '弹出框') + '</h2>' +
                    '<a class="x-close" title="点击关闭">X</a>' +
                '</div>' +
                '<div class="x-main">' +
                    '<div class="x-view">' +
                        (ops.html || '') +
                    '</div>' +
                '</div>';

            //解决ie6中 div 无法遮盖 select 的问题
            if (isIE6) {
                html += '<iframe class="x-ie-fix" style="position: absolute;top:0;left:0;z-index:-1;" scrolling="no" frameborder="0"></iframe>';
            }

            var box = createEle("div", "x-box" + (className ? " " + className : ""), html);
            Q.body.appendChild(box);

            var zIndex = ops.zIndex || 0;
            if (CURRENT_INDEX > DEF_INDEX) zIndex = Math.max(zIndex, CURRENT_INDEX);

            if (zIndex) box.style.zIndex = zIndex;

            var boxHead = getFirst(box),
                boxMain = getNext(boxHead);

            //兼容性考虑,width最好指定固定值
            if (isUNum(width)) setWidth(box, width);

            if (maxHeight) {
                if (isUNum(height) && height > maxHeight) height = maxHeight;

                if (box.offsetHeight > maxHeight) {
                    height = maxHeight;

                    addClass(box, "x-box-auto");
                }
            }

            if (isUNum(height)) {
                if (height > 50) setHeight(boxMain, height - boxHead.offsetHeight - 20);
            }

            if (boxHead.offsetWidth < 10) setWidth(boxHead, box.offsetWidth);

            if (ops.mask) self.mbox = ops.mask == "new" ? new MaskBox() : getMaskBox();

            var action_close = ops.close || "hide",
                callback_close = self.getEventCallback(action_close);

            //点击关闭事件
            self.on(".x-close", "click", action_close);

            //按ESC关闭事件
            if (ops.esc !== false) {
                self.bind(document, "keyup", function (e) {
                    if (e.keyCode == 27) callback_close();
                });
            }

            var time = ops.time;
            if (isUInt(time)) async(callback_close, time);

            self.box = box;

            fire(ops.init, self, box, ops);

            if (ops.drag !== false) {
                self.dr = setDrag(box, {
                    target: boxHead,
                    center: ops.center !== false,

                    //1.由于拖动会创建一个遮罩层,点击关闭时不会触发 .x-close 的click事件,此处检查点击元素,只有非 .x-close 元素才会执行拖动操作
                    //2.亦可将 .x-close 的click事件替换为mousedown事件以优先执行,可不必传入onCheck
                    onCheck: function (e) {
                        var target = e.target;
                        return target && target.className != "x-close";
                    }
                });
            }

            $(".x-ie-fix", box).width(box.offsetWidth - 2).height(box.offsetHeight - 2);

            //设置标题
            self.setTitle = function (title) {
                $(".x-title", boxHead).html(title);
            };
        });
    }

    //创建对话框
    function createDialogBox(ops) {
        if (!ops.width) ops.width = 320;

        var width = ops.width;

        if (ops.icon || ops.iconHtml) {
            var html =
                '<div class="fl x-ico">' +
                    (ops.iconHtml || '<img alt="" src="' + ops.icon + '"/>') +
                '</div>' +
                '<div class="fl x-dialog"' + (isUNum(width) ? ' style="width:' + (width - 60) + 'px;"' : '') + '>' + ops.html + '</div>' +
                '<div class="clear"></div>';

            ops.html = html;

            var oldInit = ops.init;

            ops.init = function (box, ops) {
                var self = this,
                    contentWidth = self.get(".x-view").offsetWidth - self.get(".x-ico").offsetWidth;

                $(".x-dialog", box).width(contentWidth);

                fire(oldInit, self, box, ops);
            };
        } else {
            ops.html = '<div class="x-dialog">' + ops.html + '</div>';
        }

        if (ops.bottom) ops.html += ops.bottom;
        if (!ops.close) ops.close = "remove";

        return WinBox(ops);
    }

    //获取底部模板代码
    function get_bottom_html(mode, style) {
        var buttonStyle = 'inline-block w-button w-' + (style || "dark"),

            html = '' +
            '<div class="x-bottom">' +
                '<div class="' + buttonStyle + ' x-submit">确定</div>' +
                (mode == 2 ? '<div class="' + buttonStyle + ' x-cancel">取消</div>' : '') +
            '</div>';

        return html;
    }

    var dialog = {
        //创建自定义弹出框
        createDialogBox: createDialogBox,

        //提示框
        alert: function (msg, fn, ops) {
            var ops = ops || {};
            if (isFunc(fn)) ops.callback = fn;
            else if (typeof fn === "object") ops = fn;

            if (!ops.title) ops.title = "提示信息";
            //ops.icon = 'images/Q/alert.gif';
            ops.iconHtml = '<div class="ico x-alert"></div>';
            ops.html = msg;

            return createDialogBox(ops);
        },
        //确认框
        confirm: function (msg, fn, ops) {
            var ops = ops || {};
            if (isFunc(fn)) ops.callback = fn;
            else if (typeof fn === "object") ops = fn;

            if (!ops.title) ops.title = "确认信息";
            ops.html = msg;
            if (!ops.bottom) ops.bottom = get_bottom_html(2);
            ops.mask = ops.mask !== false;

            var box = createDialogBox(ops);

            box.on(".x-submit", "click", "remove", true).on(".x-cancel", "click", "remove", false);

            return box;
        },
        prompt: function (msg, fn, ops) {
            var ops = ops || {};
            if (typeof fn === "object") ops = fn;

            if (!ops.title) ops.title = "输入信息";
            ops.html = '<div class="x-text">' + msg + '</div><div class="x-input"><input type="' + (ops.pwd ? 'password' : 'text') + '" /></div>';
            if (!ops.width) ops.width = 320;
            if (!ops.bottom) ops.bottom = get_bottom_html(2);

            var box = createDialogBox(ops);

            var input = box.get(".x-input>input");
            input.focus();
            input.value = def(ops.value, "");

            var submit = function () {
                var v = fire(fn, input, input.value);
                if (v !== false) box.remove();
                else setInputError(input);
            };

            //输入框快捷提交
            box.on(input, "keyup", function (e) {
                if (e.keyCode == 13) submit();
                else setInputDefault(this);
            });

            //确定与取消
            box.on(".x-submit", "click", submit).on(".x-cancel", "click", "remove");

            return box;
        },

        bottom: get_bottom_html,

        //显示加载框
        showLoading: function (ops) {
            ops = ops || { html: "正在加载数据,请稍后…" };
            //ops.icon = 'images/Q/loading.gif';
            ops.iconHtml = '<div class="ico x-loading"></div>';

            if (!ops.title) ops.title = "加载数据";

            return createDialogBox(ops);
        }
    };

    extend(Q, dialog);

    //------------------------- export -------------------------

    extend(Q, {
        getMaskBox: getMaskBox,
        maskSetup: maskSetup,

        setDrag: setDrag,

        MaskBox: MaskBox,
        DragX: DragX,

        Box: Box,
        WinBox: WinBox
    });

})();

﻿/*
* Q.UI.ContextMenu.js 多级上下文菜单(右键菜单)
* https://github.com/devin87/Q.UI.js
* author:devin87@qq.com
* update:2015/11/26 17:19
*/
(function (undefined) {
    "use strict";

    var document = window.document,

        //isObject = Q.isObject,

        fire = Q.fire,
        extend = Q.extend,
        makeArray = Q.makeArray,

        getStyle = Q.getStyle,
        getOffset = Q.offset,

        hasClass = Q.hasClass,
        addClass = Q.addClass,
        removeClass = Q.removeClass,

        createEle = Q.createEle,
        factory = Q.factory,

        E = Q.event;

    var POS_VALUE_HIDDEN = -10000;

    //设置元素位置
    function set_pos(el, x, y) {
        if (x != undefined) el.style.left = x + "px";
        if (y != undefined) el.style.top = y + "px";
    }

    //---------------------- 上下文菜单 ----------------------
    function ContextMenu(data, ops) {
        this.set(ops).init(data);
    }

    factory(ContextMenu).extend({
        //初始化菜单
        init: function (data) {
            var self = this;

            self.draw(data);

            if (self.autoHide !== false) {
                self._e0 = E.add(document, "click", function () {
                    self.hide();
                });
            }

            return self;
        },
        //菜单设置 {x,y,rangeX,rangeY}
        set: function (ops) {
            extend(this, ops, true);
            return this;
        },
        //生成所有菜单
        draw: function (data) {
            var self = this;

            //数据重置
            self._menus = [];
            self._items = [];
            self._map_menu = {};
            self._map_item = {};

            self.i = self.j = 0;
            self._active = undefined;

            //生成菜单
            var tmp = self._tmp = [],
                args;

            self.drawMenu(data);

            while ((args = tmp.shift())) {
                self.drawMenu.apply(self, args);
            }

            self._tmp = null;
            self.i = self.j = 0;

            //主菜单
            var box = self._getMenu(0).node;
            self.box = box;

            //------------ 菜单事件 ------------
            var list_box = [];
            self._menus.forEach(function (m) {
                if (m && !m.linked) list_box.push(m.node);
            });

            //移除之前绑定的事件
            if (self._e1) self._e1.off();

            self._e1 = E.add(list_box, {
                click: function (e) {
                    var el = this,
                        j = el._j;

                    if (hasClass(el, "x-disabled") || (self._getSubMenu(j) && !self.isFireAll)) return false;

                    var item = self._getItem(j);

                    fire(item.data.click, el, e, item);
                    fire(self.onclick, el, e, item);

                    self.hide();
                },
                mouseenter: function (e) {
                    var el = this,
                        i = el._i,
                        j = el._j;

                    self._hideSub(i, j);

                    addClass(el, "x-on");

                    var item = self._getItem(j);
                    fire(item.data.mouseover, el, e, item);
                    fire(self.onmouseover, el, e, item);

                    if (hasClass(el, "x-disabled")) return;

                    self._showSub(i, j, el);
                },
                mouseleave: function (e) {
                    var el = this,
                        i = el._i,
                        j = el._j,

                        sub_menu = self._getSubMenu(j);

                    if (hasClass(el, "x-disabled") || !sub_menu) removeClass(el, "x-on");
                    else if (i == 0) self._active = j;

                    var item = self._getItem(j);
                    fire(item.data.mouseout, el, e, item);
                    fire(self.onmouseout, el, e, item);
                }
            }, ".x-item");

            if (self._e2) self._e2.off();
            self._e2 = E.add(list_box, "mousedown", E.stop);

            return self.hide();
        },
        //生成菜单
        drawMenu: function (data, pid) {
            var self = this,

                list_menu = self._menus,
                list_item = self._items,
                map_menu = self._map_menu,
                map_item = self._map_item,

                box = data.box,
                is_linked_box = !!box,

                menu_index = self.i++;

            if (!box) {
                var className = data.className,
                    box = createEle("div", "x-panel" + (className ? " " + className : ""));

                box.style.width = (data.width || 120) + "px";
            }

            list_menu[menu_index] = { node: box, linked: is_linked_box, j: pid, data: data };
            if (pid) map_menu[pid] = menu_index;

            //链接菜单
            if (is_linked_box) return self;

            var global_submenu = self.subMenu !== false

            //循环生成菜单项
            for (var i = 0, data_items = data.items; i < data_items.length; i++) {
                var m = data_items[i];
                if (!m) continue;

                var item = document.createElement("div"),
                    item_index = self.j++;

                item._i = menu_index;
                item._j = item_index;

                list_item[item_index] = { node: item, i: menu_index, j: item_index, data: m };
                if (m.id != undefined) map_item[m.id] = item_index;

                if (m.split) {
                    item.className = "x-split";
                } else {
                    item.className = "x-item" + (m.disabled ? " x-disabled" : "");
                    item.x = m.x != undefined ? m.x : ""; //自定义值

                    var group = m.group,
                        has_submenu = global_submenu && group && (group.box || (group.items && group.items.length > 0)),
                        ico = m.ico;

                    var html = m.html ||
                        '<div class="x-icon">' +
                            (ico ? /^<.+>$/.test(ico) ? ico : '<img alt="" src="' + ico + '">' : '') +
                        '</div>' +
                        '<div class="x-text"' + (m.title ? ' title="' + m.title + '"' : '') + '>' +
                            m.text +
                        '</div>' +
                        (has_submenu ? '<div class="arrow"></div>' : '');

                    item.innerHTML = html;

                    if (has_submenu) self._tmp.push([group, item_index]);
                }

                box.appendChild(item);
            }

            Q.body.appendChild(box);

            return self;
        },

        //根据菜单索引获取菜单
        _getMenu: function (i) {
            return this._menus[i];
        },

        //根据菜单项索引获取子菜单
        _getSubMenu: function (j) {
            var i = this._map_menu[j];
            return i != undefined ? this._getMenu(i) : undefined;
        },

        //根据菜单项索引获取菜单项
        _getItem: function (j) {
            return this._items[j];
        },

        //根据指定的id获取菜单项
        getItem: function (id) {
            var j = this._map_item[id];

            return j != undefined ? this._items[j] : undefined;
        },

        //更改菜单项的值
        setItemText: function (id, text) {
            var item = this.getItem(id);
            if (!item) return;

            item.data.text = text;

            var node_text = item.node.childNodes[1];
            if (node_text) node_text.innerHTML = text;
        },

        //处理菜单项
        processItems: function (ids, fn) {
            var self = this;

            makeArray(ids).forEach(function (id) {
                var item = self.getItem(id);
                if (item) fn(item.node, item);
            });

            return self;
        },

        //处理菜单项(className)
        _processItems: function (ids, fn, className) {
            return this.processItems(ids, function (node) {
                fn(node, className);
            });
        },

        //启用菜单项
        enableItems: function (ids) {
            return this._processItems(ids, removeClass, "x-disabled");
        },
        //禁用菜单项
        disableItems: function (ids) {
            return this._processItems(ids, addClass, "x-disabled");
        },
        //显示菜单项
        showItems: function (ids) {
            return this._processItems(ids, removeClass, "hide");
        },
        //隐藏菜单项
        hideItems: function (ids) {
            return this._processItems(ids, addClass, "hide");
        },

        //元素智能定位
        _setPos: function (menu, x, y, pbox) {
            var self = this,
                rangeX = self.rangeX,
                rangeY = self.rangeY,

                box = menu.node,
                width = box.offsetWidth,
                height = box.offsetHeight,

                data = menu.data || {},
                maxHeight = data.maxHeight;

            //如果菜单高度超出最大高度,则启用垂直滚动条
            if (maxHeight) {
                var realHeight = Math.max(height, box.scrollHeight),
                    hasScrollbar = realHeight > maxHeight;

                box.style.height = hasScrollbar ? maxHeight + "px" : "auto";

                //是否固定宽度,默认菜单宽度会增加17px
                if (!self.fixedWidth) box.style.width = (data.width + (hasScrollbar ? 17 : 0)) + "px";

                height = box.offsetHeight;
            }

            if (x == undefined) x = self.x || 0;
            if (y == undefined) y = self.y || 0;

            if (rangeX && x + width > rangeX) {
                x = rangeX - width;
                if (pbox) x = x - pbox.offsetWidth + 3;
            }

            if (rangeY && y + height > rangeY) y = rangeY - height - 1;

            if (x < 0) x = 0;
            if (y < 0) y = 0;

            set_pos(box, x, y);

            self.x = x;
            self.y = y;

            return self;
        },

        //显示子菜单(如果有)
        _showSub: function (i, j, el) {
            var self = this,

                menu = self._getMenu(i),
                sub_menu = self._getSubMenu(j);

            if (sub_menu) {
                var node = menu.node,
                    sub_node = sub_menu.node,
                    offset = getOffset(el),

                    zIndex = +getStyle(node, "zIndex"),
                    zIndex_sub = +getStyle(sub_node, "zIndex");

                if (zIndex_sub <= zIndex) sub_node.style.zIndex = zIndex + 1;

                self._setPos(sub_menu, offset.left + el.offsetWidth - 2, offset.top, node);
            }

            self.i = i;
            self.j = j;

            return self;
        },

        //取消高亮的项
        _inactive: function (j) {
            var item = this._getItem(j);
            if (item) removeClass(item.node, "x-on");

            return this;
        },

        //隐藏子菜单
        _hideSub: function (i, j) {
            var self = this;

            if (i <= self.i && j != self.j) {
                var list_menu = self._menus, k;

                //隐藏子一级菜单(当前菜单以后的菜单)
                for (k = list_menu.length - 1; k > i; k--) {
                    self._hide(list_menu[k]);
                }

                //移除子一级菜单高亮的项
                if (i < self.i) {
                    for (k = self.i; k > i; k--) {
                        var p_menu = self._getMenu(k);
                        self._inactive(p_menu.j);
                    }
                }
            }

            //移除同一菜单里上次高亮的项(有子菜单的项)
            if (i == self.i && j != self.j && self._getSubMenu(self.j)) self._inactive(self.j);

            return self;
        },
        //隐藏指定菜单
        _hide: function (menu) {
            if (menu) {
                var node = menu.node;
                if (node.style.left != POS_VALUE_HIDDEN) set_pos(node, POS_VALUE_HIDDEN, POS_VALUE_HIDDEN);
            }

            return this;
        },

        //显示
        show: function (x, y) {
            return this._setPos(this._menus[0], x, y);
        },
        //隐藏
        hide: function () {
            var self = this;

            self._menus.forEach(function (m) {
                self._hide(m);
            });

            if (self._active != undefined) {
                self._inactive(self._active);
                self._active = undefined;
            }

            return self;
        },

        //菜单是否为隐藏状态
        isHidden: function () {
            var self = this,
                box = self.box,
                x = parseFloat(box.style.left),
                y = parseFloat(box.style.top);

            return x <= -box.offsetWidth || y <= -box.offsetHeight;
        },

        //自动切换显示或隐藏
        toggle: function (x, y) {
            return this.isHidden() ? this.show(x, y) : this.hide();
        },

        //注销菜单
        destroy: function () {
            var self = this;

            if (self._e0) self._e0.off();
            if (self._e1) self._e1.off();

            self._menus.forEach(function (menu) {
                menu.node.innerHTML = '';
                $(menu.node).remove();
            });

            Object.forEach(self, function (prop) {
                self[prop] = null;
            });

            return self;
        }
    });

    //------------------------- export -------------------------

    Q.ContextMenu = ContextMenu;

})();

﻿/*
* Q.UI.DropdownList.js 下拉列表
* https://github.com/devin87/Q.UI.js
* author:devin87@qq.com
* update:2015/11/26 13:04
*/
(function (undefined) {
    "use strict";

    var document = window.document,

        def = Q.def,
        fire = Q.fire,

        isObject = Q.isObject,

        getFirst = Q.getFirst,
        getLast = Q.getLast,

        getWidth = Q.width,
        setWidth = getWidth,

        hasClass = Q.hasClass,
        addClass = Q.addClass,
        removeClass = Q.removeClass,

        createEle = Q.createEle,
        factory = Q.factory,

        ie = Q.ie,

        E = Q.event;

    //---------------------- 下拉列表 ----------------------

    function DropdownList(ops) {
        ops = ops || {};

        var self = this;

        self.box = ops.box;
        self.items = ops.items || [];
        self.multiple = ops.multiple;

        //默认值或索引
        self.value = ops.value;
        self.index = ops.index || 0;

        self.ops = ops;
    }

    factory(DropdownList).extend({
        init: function () {
            var self = this,

                ops = self.ops,
                box = self.box,

                isDropdownList = !self.multiple;

            var html =
                (isDropdownList ?
                '<div class="x-sel-tag">' +
                    '<div class="x-sel-text"></div>' +
                    '<div class="x-sel-arrow">' +
                        '<div class="arrow arrow-down"></div>' +
                    '</div>' +
                '</div>' : '') +
                '<div class="x-panel x-sel-list"></div>';

            addClass(box, "x-sel");

            box.innerHTML = html;

            var elList = getLast(box),
                elTag, elText, elArrow;

            setWidth(elList, ops.width || box.offsetWidth - 2);

            if (isDropdownList) {
                elTag = getFirst(box);
                elText = getFirst(elTag);
                elArrow = getLast(elTag);

                self.elText = elText;
                self.elArrow = elArrow;

            }

            self.elList = elList;

            //------------- Event --------------

            //item 相关事件
            var listener_item;

            //下拉列表
            if (isDropdownList) {
                if (ops.autoHide !== false) {
                    E.add(document, "mousedown", function () {
                        self.hide();
                    });
                }

                E.add(box, "mousedown", function (e) {
                    self.toggle();

                    return false;
                });

                listener_item = {
                    mousedown: E.stop,
                    mouseup: function (e) {
                        var index = this.x,
                            item = self.items[index];

                        fire(self.onclick, self, item, index);

                        if (item.disabled) return;

                        self.hide();
                        if (index != self.index) self.select(index);
                    },
                    mouseenter: function () {
                        var index = this.x,
                            item = self.items[index];

                        if (!item.disabled) self.active(index);
                    }
                };
            } else {
                self.selectedItems = [];
                self.seletedMap = {};

                //多选框事件
                listener_item = {
                    mousedown: function (e) {
                        var el = this;
                        if (hasClass(el, "x-disabled")) return;

                        var index = this.x,
                            //items = self.items,
                            firstItem = self.selectedItems[0],

                            shiftKey = e.shiftKey,
                            ctrlKey = e.ctrlKey;

                        if (shiftKey || !ctrlKey) self.clearSelect();

                        if (shiftKey) {
                            var start = firstItem ? firstItem.index : index,
                                end = index;

                            if (start > end) {
                                end = start;
                                start = index;
                            }
                            for (var i = start; i <= end; i++) {
                                self.active(i);
                            }
                        } else {
                            if (ctrlKey && self.seletedMap[index]) self.inactive(index);
                            else self.select(index);
                        }
                    }
                };

                if (ie < 10) listener_item.selectstart = E.stop;
            }

            E.add(elList, listener_item, ".x-item");

            return self.draw();
        },
        //绘制下拉视图
        draw: function () {
            var self = this,
                ops = self.ops,
                items = self.items,
                elList = self.elList,
                hasTitle = ops.hasTitle,
                maxHeight = ops.maxHeight,
                map = {};

            elList.innerHTML = "";

            items.forEach(function (m, i) {
                m.index = i;
                map[m.value] = m;

                var text = m.text || '',
                    node = createEle("div", "x-item" + (m.group ? " x-sel-group" : m.disabled ? " x-disabled" : ""), text);

                if (hasTitle) node.title = m.title || text.toText();

                node.x = i;
                m.node = node;

                elList.appendChild(node);
            });

            self.map = map;

            if (maxHeight) {
                var realHeight = elList.offsetHeight;
                if (realHeight > maxHeight) elList.style.height = maxHeight;
            }

            var value = self.value,
                index = self.index;

            if (value) {
                var item = self.find(value);
                if (item) index = item.index;
            }

            self.select(index);

            return self.multiple ? self : self.hide();
        },

        //添加下拉项,类似option
        add: function (text, value, title) {
            this.items.push(isObject(text) ? text : { text: text, value: value, title: title });
            return this;
        },
        //查找指定值的项 eg:{text,value,index}
        find: function (value) {
            return this.map[value];
        },
        //清除选中项
        clearSelect: function () {
            var self = this;

            var inactive = function (item) {
                if (item.node) removeClass(item.node, "selected");
            };

            //多选框
            if (self.multiple) {
                self.selectedItems.forEach(inactive);

                self.selectedItems = [];
                self.seletedMap = {};
            } else {
                inactive({ node: self._el });
            }

            return self;
        },
        //设置项高亮
        active: function (index) {
            var self = this,
                item = self.items[index];

            if (!item) return self;

            var node = item.node,
                flag;

            //多选框
            if (self.multiple) {
                if (!item.disabled && !self.seletedMap[index]) {
                    self.selectedItems.push(item);
                    self.seletedMap[index] = true;

                    flag = true;
                }
            } else {
                self.clearSelect();
                flag = true;

                self._el = node;
            }

            if (flag) addClass(node, "selected");

            return this;
        },
        //取消项高亮
        inactive: function (index) {
            var self = this,
                item = self.items[index];

            if (self.multiple) {
                self.selectedItems = self.selectedItems.filter(function (item) {
                    return item.index != index;
                });
                self.seletedMap[index] = undefined;
            }

            removeClass(item.node, "selected");

            return self;
        },
        //选中指定索引的项
        //isNotChange:是否不触发change事件
        select: function (index, isNotChange) {
            var self = this,
                items = self.items,
                item = items[index];

            if (!item) return self;

            self.text = def(item.text, "");
            self.value = def(item.value, "");

            if (self.elText) self.elText.innerHTML = self.text;

            //多选框
            self.active(index);

            if (!isNotChange && index != self.index) fire(self.onchange, self, item, index);

            self.index = index;

            return self;
        },

        //显示下拉列表
        show: function () {
            this.elList.style.display = "";
            return this.select(this.index);
        },
        //隐藏下拉列表
        hide: function () {
            this.elList.style.display = "none";
            return this;
        },
        //自动切换显示或隐藏下拉列表
        toggle: function () {
            return this.elList.style.display == "none" ? this.show() : this.hide();
        }
    });

    //------------------------- export -------------------------

    Q.DropdownList = DropdownList;

})();

﻿/*
* Q.UI.DataPager.js 数据分页
* https://github.com/devin87/Q.UI.js
* author:devin87@qq.com
* update:2015/10/16 10:26
*/
(function (undefined) {
    "use strict";

    var factory = Q.factory;

    //---------------------- 数据分页 ----------------------

    //生成分页页码数据 eg:create_pager_bar_data(9,28,5) => [1,0,4,5,6,7,8,0,28] 0表示省略号，其它表示页码
    //size:要显示的页码数量(包括省略号),不小于7的奇数
    //totalPage:总页数
    //page:当前页索引
    function create_pager_bar_data(size, totalPage, page) {
        var tmp = [1];

        //只有一页的情况
        if (totalPage == 1) return tmp;

        //if (size < 7) size = 7;
        //else if (size % 2 == 0) size++;

        var first = 2, last;

        //总页数小于或等于页码数量
        if (totalPage <= size) {
            last = totalPage;
        } else {
            var sideCount = (size - 5) / 2;

            //左边无省略号 eg: [1,2,3,4,5,0,20]
            if (page <= sideCount + 3) {
                last = size - 2;
            } else if (page > totalPage - sideCount - 3) {
                //右边无省略号 eg: [1,0,16,17,18,19,20]
                first = totalPage - size + 3;  //=last-(size-4)
                last = totalPage - 1;
            } else {
                //左右均有省略号 eg: [1,0,8,9,10,0,20]
                first = page - ~~sideCount;
                last = first + size - 5;
            }
        }

        if (first > 2) tmp.push(0);
        for (var i = first; i <= last; i++) tmp.push(i);

        if (totalPage > size) {
            if (last < totalPage - 1) tmp.push(0);
            tmp.push(totalPage);
        }

        return tmp;
    }

    //构造器:分页对象
    function DataPager(ops) {
        var self = this;

        self.ops = ops;
        self.size = ops.size || 9;
        self.cache = {};

        if (ops.href == undefined) {
            var boxNav = ops.boxNav;
            if (boxNav) {
                $(boxNav).on("click", "li", function () {
                    self.go($(this).attr("x"), true);

                    if (self.onclick) self.onclick.call(self, self.page);
                });
            }
        } else {
            ops.cache = ops.preload = false;
        }

        self.set(ops.totalCount, ops.pageSize || 10).setData(ops.data).go(ops.page);
    }

    factory(DataPager).extend({
        //设置总的数据条数和每页显示的数据条数
        set: function (totalCount, pageSize) {
            var self = this;
            if (totalCount != undefined) self.totalCount = totalCount;
            if (pageSize != undefined) self.pageSize = pageSize;

            if (self.totalCount != undefined) self.totalPage = Math.ceil(self.totalCount / self.pageSize);

            return self;
        },
        //设置数据列表(为传入ops.load)
        //data:数据列表
        setData: function (data) {
            if (data) {
                this.data = data;
                this.set(data.length);
            }
            return this;
        },
        //加载数据
        _load: function (page, callback) {
            var self = this,
                ops = self.ops,
                onload = self.onload;

            self.page = page;

            var complete = function (list) {
                if (onload) onload.call(self, list);
                if (callback) callback.call(self, list);

                return self;
            };

            //本地数据分页
            if (!ops.load) {
                var pageSize = self.pageSize,
                    start = (page - 1) * pageSize;

                return complete(self.data.slice(start, start + pageSize));
            }

            //使用缓存
            var list = ops.cache ? self.cache[page] : undefined;
            if (list) return complete(list);

            //获取远程数据
            ops.load(page, function (data) {
                list = data.data || [];
                if (ops.cache) self.cache[page] = list;

                self.set(data.totalCount, data.pageSize);

                complete(list);
            });

            return self;
        },
        //加载并渲染数据
        load: function (page) {
            return this._load(page, this.draw);
        },
        //重新载入当前数据并渲染
        reload: function (page) {
            return this.load(page || this.page);
        },
        //跳转到指定页
        //forced:是否强制跳转
        go: function (page, forced) {
            var self = this;

            if (isNaN(page)) return self;

            page = +page;
            if (self.totalPage != undefined && page > self.totalPage) page = self.totalPage;
            if (page < 1) page = 1;

            if (page == self.page && !forced) return self;

            self.load(page);
            if (self.ops.load && self.ops.preload) self._load(page + 1);

            if (self.onchange) self.onchange.call(self, page);

            return self;
        },
        //绘制导航区(分页按钮)
        drawNav: function () {
            var self = this,
                ops = self.ops,
                boxNav = ops.boxNav;

            if (!boxNav) return self;

            var totalCount = self.totalCount,
                pageSize = self.pageSize,

                totalPage = self.totalPage,
                page = self.page,

                href = ops.href,
                text = ops.text || {},

                list_bar = create_pager_bar_data(self.size, totalPage, page);

            if (href != undefined) href += href.indexOf("?") != -1 ? "&" : "?";

            var get_html_bar = function (i, title, className) {
                if (!className) className = i == page ? "on" : (i ? "" : "skip");

                return '<li' + (className ? ' class="' + className + '"' : '') + ' x="' + i + '">' + (href ? '<a' + (i ? ' href="' + href + 'page=' + i + '"' : '') + '>' + title + '</a>' : title) + '</li>';
            };

            var draw_size = ops.drawSize || function (self, html_page, html_count) {
                return html_page + '/' + html_count;
            };

            var html =
                '<div class="inline-block pager-bar' + (href ? ' pager-link' : '') + '">' +
                    '<ul>' +
                        get_html_bar(Math.max(page - 1, 1), text.prev || '&lt;上一页', "prev") +
                        list_bar.map(function (i) {
                            return get_html_bar(i, i || "…");
                        }).join('') +
                        get_html_bar(Math.min(page + 1, totalPage), text.next || '下一页&gt;', "next") +
                    '</ul>' +
                '</div>' +
                (ops.showSize !== false ?
                '<div class="inline-block pager-count">' +
                    draw_size(self, '每页<span class="page-size">' + pageSize + '</span>条', '共<span class="total-count">' + totalCount + '</span>条数据') +
                '</div>' : '');

            $(boxNav).html(html);
        },
        //绘制UI
        draw: function (list) {
            this.drawNav();
            this.ops.draw(list);

            return this;
        }
    });

    //------------------------- export -------------------------

    Q.DataPager = DataPager;

})();

﻿/*
* Q.UI.ColorPicker.js 颜色选择器
* https://github.com/devin87/Q.UI.js
* author:devin87@qq.com
* update:2015/12/08 14:39
*/
(function (undefined) {
    "use strict";

    var document = window.document,

        isFunc = Q.isFunc,

        extend = Q.extend,
        //fire = Q.fire,

        getFirst = Q.getFirst,
        getNext = Q.getNext,
        getLast = Q.getLast,

        //getOffset = Q.offset,

        createEle = Q.createEle,
        factory = Q.factory,

        E = Q.event;

    var POS_VALUE_HIDDEN = -10000;

    //设置元素位置
    function set_pos(el, x, y) {
        if (x != undefined) el.style.left = x + "px";
        if (y != undefined) el.style.top = y + "px";
    }

    //---------------------- util ----------------------

    var RE_RGB = /^rgb\((\d+),(\d+),(\d+)\)$/i,
        RE_Hex = /^#?([0-9A-F]{3}|[0-9A-F]{6})$/i;

    //int值转为16进制颜色
    function int2hex(n) {
        return "#" + ("00000" + n.toString(16)).slice(-6);   //faster

        //var str = ("00000" + n.toString(16)).slice(-6);
        //return "#" + str.slice(4) + str.slice(2, 4) + str.slice(0, 2);
    }

    //RGB颜色转为16进制颜色
    function rgb2hex(r, g, b) {
        return int2hex(r * 65536 + g * 256 + b);
        //return int2hex(r + g * 256 + b * 65536);
    }

    //转为16进制颜色
    function toHex(color) {
        if (typeof color == "number") return int2hex(color);

        color = color.replace(/\s+/g, "");
        if (!RE_RGB.test(color)) return color;

        return rgb2hex(+RegExp.$1, +RegExp.$2, +RegExp.$3);
    }

    //转为RGB颜色
    function toRGB(color) {
        if (typeof color == "number") color = int2hex(color);
        if (!RE_Hex.test(color)) return color;

        color = RegExp.$1;

        if (color.length == 3) {
            var a = color.charAt(0), b = color.charAt(1), c = color.charAt(2);
            color = a + a + b + b + c + c;
        }

        return "rgb(" + parseInt(color.substr(0, 2), 16) + "," + parseInt(color.substr(2, 2), 16) + "," + parseInt(color.substr(4, 2), 16) + ")";
    }

    //---------------------- ColorPicker ----------------------

    var LIST_COLOR = ['#000000', '#333333', '#666666', '#999999', '#cccccc', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#00ffff', '#ff00ff'];

    //颜色选择器
    function ColorPicker(ops) {
        var self = this;

        //默认行数与列数
        self.row = 12;  //不得超过 LIST_COLOR.length
        self.col = 21;

        self.set(ops).init();
    }

    factory(ColorPicker).extend({
        //初始化
        init: function () {
            var self = this;

            var html =
                '<div class="xp-title">' +
                    '<div class="xp-preview"></div>' +
                    '<div class="xp-val"></div>' +
                    '<div class="xp-type">' +
                        '<select>' +
                            '<option value="Cube" selected="selected">立方色</option>' +
                            '<option value="Series">连续色调</option>' +
                            '<option value="Gray">灰度等级</option>' +
                        '</select>' +
                    '</div>' +
                '</div>' +
                '<div class="xp-table">' +
                    '<table>' +
                        ('<tr>' + '<td></td>'.repeat(self.col) + '</tr>').repeat(self.row) +
                    '</table>' +
                '</div>';

            var box = createEle("div", "x-picker", html),

                boxTitle = getFirst(box),
                boxPreview = getFirst(boxTitle),
                boxValue = getNext(boxPreview),
                boxType = getLast(boxTitle),

                boxTable = getLast(box),
                table = getFirst(boxTable);

            Q.body.appendChild(box);

            self.box = box;
            self.table = table;
            self.boxPreview = boxPreview;
            self.boxValue = boxValue;

            //------------------- init event -------------------

            //类型切换
            E.add(getFirst(boxType), "change", function () {
                self["draw" + this.value + "Color"]();
            });

            E.add(table, {
                //鼠标移动,预览颜色
                mouseover: function (e) {
                    self.setPreview(this.style.backgroundColor);
                },

                //选择并设置颜色
                click: function (e) {
                    var color = this.style.backgroundColor;
                    self.fire(color).hide();
                }
            }, "td");

            E.add(box, "click", E.stop);
            E.add(document, "click", function (e) { self.hide(); });

            return self.drawCubeColor().hide();
        },
        //设置 { isHex,callback }
        //isHex:输出为16进制颜色
        set: function (ops) {
            extend(this, ops, true);

            return this.setPreview((ops || {}).color);
        },

        //触发回调函数
        fire: function (color) {
            var self = this;
            if (isFunc(self.callback)) self.callback.call(self, self.isHex ? toHex(color) : color);
            return self;
        },
        //设置预览颜色
        setPreview: function (color) {
            var self = this;

            if (color) {
                self.boxPreview.style.backgroundColor = color;
                self.boxValue.innerHTML = toHex(color).toUpperCase();
            }

            return self;
        },

        //填充单元格颜色
        fillColor: function (i, j, color) {
            this.table.rows[i].cells[j].style.backgroundColor = color;
            return this;
        },

        //画左边的颜色
        drawLeftColor: function () {
            var self = this,
                row = self.row;

            for (var i = 0; i < row; i++) {
                self.fillColor(i, 0, "#000").fillColor(i, 1, LIST_COLOR[i]).fillColor(i, 2, "#000");
            }

            return self;
        },
        //画立方色
        drawCubeColor: function () {
            var self = this,
                row = self.row,
                col = self.col,
                start = 0,
                step = 0x330000,
                color;

            self.drawLeftColor();

            for (var i = 0; i < row; i++) {
                if (i > 5) color = start = 0x990000 + (i - 6) * 0x000033;
                else color = start = 0x0 + i * 0x000033;

                for (var j = 3; j < col; j++) {
                    self.fillColor(i, j, int2hex(color));
                    color += 0x003300;
                    if ((j - 2) % 6 == 0) start += step, color = start;
                }
            }

            return self;
        },
        //画连续色
        drawSeriesColor: function () {
            var self = this,
                row = self.row,
                col = self.col,
                start = 0xCCFFFF,
                step = 0x660000,
                flag = 1,
                color;

            self.drawLeftColor();

            for (var i = 0; i < row; i++) {
                if (i > 5) color = start = 0xFF00FF + (i - 6) * 0x003300;
                else color = start = 0xCCFFFF - i * 0x003300;

                flag = 1;

                for (var j = 3; j < col; j++) {
                    self.fillColor(i, j, int2hex(color));
                    color -= 0x000033 * flag;
                    if ((j - 2) % 6 == 0) {
                        flag *= -1;
                        start -= step;
                        color = start - ((flag > 0) ? 0 : 0x0000FF);
                    }
                }
            }

            return self;
        },
        //画灰度等级色
        drawGrayColor: function () {
            var self = this,
                row = self.row,
                col = self.col,
                color = 0xffffff;

            for (var i = 0; i < row; i++) {
                for (var j = 0; j < col; j++) {
                    self.fillColor(i, j, int2hex(color));

                    if (color <= 0) color = 0x000000;
                    else color -= 0x010101;
                }
            }

            return self;
        },

        //显示
        show: function (x, y) {
            var self = this;

            if (x == undefined) x = self.x;
            else self.x = x;

            if (y == undefined) y = self.y;
            else self.y = y;

            set_pos(self.box, x, y);

            return self;
        },
        //隐藏
        hide: function () {
            set_pos(this.box, POS_VALUE_HIDDEN, POS_VALUE_HIDDEN);
            return this;
        },

        //是否为隐藏状态
        isHidden: function () {
            var self = this,
                box = self.box,
                x = parseFloat(box.style.left),
                y = parseFloat(box.style.top);

            return x <= -box.offsetWidth || y <= -box.offsetHeight;
        },

        //自动切换显示或隐藏
        toggle: function (x, y) {
            return this.isHidden() ? this.show(x, y) : this.hide();
        }
    });

    //------------------------- export -------------------------

    Q.toHex = toHex;
    Q.toRGB = toRGB;
    Q.ColorPicker = ColorPicker;

})();