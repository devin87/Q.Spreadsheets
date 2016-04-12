/// <reference path="../lib/Q.js" />
/// <reference path="../lib/Q.UI.js" />
/*
* util.js
* author:devin87@qq.com
* update: 2016/04/11 15:35
*/
(function (window, undefined) {
    "use strict";

    var isObject = Q.isObject,

        extend = Q.extend,

        createEle = Q.createEle,
        parseHTML = Q.parseHTML,

        getFirst = Q.getFirst,

        toMap = Q.toMap;

    //----------------------- util -----------------------
    //弹出错误信息,text会经过html编码
    function alertError(text, fn, ops) {
        if (isObject(fn)) ops = fn;
        if (!ops) ops = {};

        ops.width = 460;

        return Q.alert('<span class="hot">' + text.htmlEncode() + '</span>', ops);
    }

    //将数组或字符串的每一个字符转为map映射(值为true)
    function toSet(list_or_str) {
        return toMap(typeof list_or_str == "string" ? list_or_str.split("") : list_or_str, true);
    }

    //浮点数运算通用函数
    function Num_float_op(n1, n2, op) {
        var str1 = n1.toString(), str2 = n2.toString(), index1 = str1.indexOf("."), index2 = str2.indexOf("."), r1 = 0, r2 = 0, m;
        if (index1 != -1) r1 = str1.length - index1 - 1;
        if (index2 != -1) r2 = str2.length - index2 - 1;

        m = Math.pow(10, Math.max(r1, r2));

        var fix_n1 = Math.round(n1 * m), fix_n2 = Math.round(n2 * m);

        //将浮点数转换为整数后操作
        switch (op) {
            case 1: return (fix_n1 + fix_n2) / m;       //加法
            case 2: return (fix_n1 - fix_n2) / m;       //减法,可省略(=Num_float_op(n1,-n2,1))
            case 3: return fix_n1 * fix_n2 / (m * m);   //乘法
            case 4: return fix_n1 / fix_n2;             //除法
        }
    }

    //二分法查找,返回最接近且小于或等于该数值的索引 eg:([9,30,36,42],38) => 2
    //若返回-2,则表示list不存在或n小于list[0]
    //若返回-1,则表示n>list[list.length-1]
    //list:排序好的number数组
    function findInList(list, n) {
        var len = list.length, last = len - 1;

        if (len == 0 || n < list[0]) return -2;
        if (n > list[last]) return -1;

        if (n == list[0]) return 0;
        if (n == list[last]) return last;

        var left = 0, right = last, i;
        while (left < right) {
            i = ~~((left + right) / 2);  //~~相当于Math.floor

            if (list[i] > n) {
                right = i;
            } else {
                if (list[i + 1] == n) return i + 1;
                if (list[i + 1] > n) return i;
                left = i;
            }
        }
    }

    //----------------------- MeasureText -----------------------

    var elMeasureText, measureText;

    //文字测量对象
    function MeasureText(font) {
        if (!elMeasureText) {
            var div = parseHTML('<div style="position:absolute;left:-100000px;top:-100000px;"><span></span></div>');
            elMeasureText = getFirst(div);
            Q.body.appendChild(div);
        }

        this.el = elMeasureText;

        this.reset(font);
    }

    MeasureText.prototype = {
        //设置渲染样式 eg:bold、italic
        css: function (key, value) {
            var result = $(this.el).css(key, value);
            return value != undefined ? this : result;
        },
        //设置字体
        setFont: function (font) {
            if (font) this.css("font", font);
            return this;
        },
        //重置样式
        reset: function (font) {
            this.el.style.cssText = "white-space:nowrap;";
            return this.setFont(font);
        },
        //获取结果
        //type:text|html
        getResult: function (str, type) {
            var el = this.el;

            $(el)[type || "text"](str);

            var width = el.offsetWidth, height = el.offsetHeight;
            el.innerHTML = "";

            return { width: width, height: height };
        },
        //获取文本宽度
        //type:text|html
        getWidth: function (str, type) {
            return this.getResult(str, type).width;
        },
        //获取文本高度
        //type:text|html
        getHeight: function (str, type) {
            return this.getResult(str, type).height;
        },
        //销毁对象
        destroy: function () {
            $(this.el).remove();
            this.el = elMeasureText = null;
        }
    };

    //获取文字测量对象
    function getMeasureText(font) {
        if (measureText) return measureText.reset(font);

        measureText = new MeasureText(font);

        return measureText;
    }

    //处理菜单宽度
    function processMenuWidth(data) {
        var maxTextWidth = (data.width || 120) - 50;
        data.items.forEach(function (item) {
            var width = measureText.getWidth(item.text), group = item.group;
            if (width > maxTextWidth) maxTextWidth = width;

            if (group && group.items) processMenuWidth(group);
        });

        data.width = maxTextWidth + 50;

        return data;
    }

    //----------------------- export -----------------------

    extend(Q, {
        alertError: alertError,
        toSet: toSet,

        Num: {
            //浮点数加法修正 eg:1.1 + 0.1
            add: function (n1, n2) {
                return Num_float_op(n1, n2, 1);
            },
            //浮点数减法修正 eg:1.1 - 1
            sub: function (n1, n2) {
                return Num_float_op(n1, n2, 2);
            },
            //浮点数乘法修正 eg:1.4 * 7
            mul: function (n1, n2) {
                return Num_float_op(n1, n2, 3);
            },
            //浮点数除法修正 
            //eg:1.2 / 0.1 
            //eg:1.2 / 3
            div: function (n1, n2) {
                return Num_float_op(n1, n2, 4);
            }
        },

        findInList: findInList,

        getMeasureText: getMeasureText,
        processMenuWidth: processMenuWidth,

        MeasureText: MeasureText
    });

})(window);