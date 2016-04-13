/* Web Spreadsheets, https://github.com/devin87/Q.Spreadsheets */
﻿/*
* init.js
* author:devin87@qq.com
* update: 2015/11/19 17:25
*/
(function (window, undefined) {
    "use strict";
})(window);

﻿/*
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

﻿/*
* core.js
* author:devin87@qq.com
* update: 2015/11/26 14:34
*/
(function (window, undefined) {
    "use strict";

    var toMap = Q.toMap;

    //----------------------- ColName & ColIndex -----------------------

    //列头名称及映射
    var LIST_COLHEAD_LETTERS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"],
        MAP_COLHEAD_LETTERS = toMap(LIST_COLHEAD_LETTERS),

        COLHEAD_LEN1 = 26,
        COLHEAD_LEN2_LEVEL = COLHEAD_LEN1 * COLHEAD_LEN1,
        COLHEAD_LEN2 = COLHEAD_LEN2_LEVEL + COLHEAD_LEN1,
        COLHEAD_LEN3 = COLHEAD_LEN2 * COLHEAD_LEN1 + COLHEAD_LEN1;

    //根据列索引获取列头名称(0<=index<18278)
    function getColName(index) {
        if (index < COLHEAD_LEN1) return LIST_COLHEAD_LETTERS[index];
        if (index < COLHEAD_LEN2) return LIST_COLHEAD_LETTERS[Math.floor(index / COLHEAD_LEN1) - 1] + LIST_COLHEAD_LETTERS[index % COLHEAD_LEN1];
        if (index < COLHEAD_LEN3) {
            var i = Math.floor((index - COLHEAD_LEN2) / COLHEAD_LEN2_LEVEL), j = Math.floor((index - (i + 1) * COLHEAD_LEN2_LEVEL - COLHEAD_LEN1) / COLHEAD_LEN1);
            return LIST_COLHEAD_LETTERS[i] + LIST_COLHEAD_LETTERS[j] + LIST_COLHEAD_LETTERS[index % COLHEAD_LEN1];
        }
    }

    //范围更广(n>=0),但性能略有降低,且Excel 2010也只支持16384列(XFD列),不会超过18278,所以上面函数足以
    //function getColName2(n) {
    //    var x = 0, tmp = [COLHEAD_LEN1], name = "";

    //    while (tmp[x] <= n) {
    //        tmp.push(Math.pow(COLHEAD_LEN1, ++x + 1) + tmp[x - 1]);
    //    }

    //    while (--x >= 0) {
    //        var pow = Math.pow(COLHEAD_LEN1, x + 1), i = Math.floor((n - tmp[x]) / pow);
    //        name += LIST_COLHEAD_LETTERS[i];
    //        n -= (i + 1) * pow;
    //    }
    //    name += LIST_COLHEAD_LETTERS[n % COLHEAD_LEN1];

    //    return name;
    //}

    //根据列头名称获取列索引
    function getColIndex(name) {
        name = name.toUpperCase();
        var x = name.length, k = -1, index = 0;
        while (--x >= 0) {
            var i = MAP_COLHEAD_LETTERS[name.charAt(++k)] + 1;
            index += i * Math.pow(COLHEAD_LEN1, x);
        }
        return index - 1;
    }

    Q.SS = {
        getColName: getColName,
        getColIndex: getColIndex
    };
})(window);

﻿/*
* events.js
* author:devin87@qq.com
* update: 2016/01/15 12:58
*/
(function () {
    "use strict";

    Q.SS.Events = {
        LOAD: "load",

        WORKBOOK_READY: "workbook:ready",

        SHEET_BEFORE: "sheet:before",
        SHEET_READY: "sheet:ready",
        SHEET_LOAD: "sheet:load",

        UI_RESIZE: "ui:resize"
    };

})();

﻿/*
* errors.js
* author:devin87@qq.com
* update: 2016/01/15 12:58
*/
(function () {
    "use strict";

    Q.SS.Errors = {
        PARAMETER_ERROR: -2,
        SHEET_NOT_FOUND: -3,
        SHEET_ALREADY_EXISTS: -4,
        SHEET_NAME_NOT_FOUND: -5,
        SHEET_NAME_ALREADY_EXISTS: -6,
        SHEET_INDEX_IS_SAME: -7,
        SHEET_NAME_IS_SAME: -8,
        SHEET_NUMBER_LESS_TWO: -9
    };
})();

﻿/*
* cell-format.parse.js
* author:devin87@qq.com & zhangxin
* update: 2015/11/26 14:33
*/
(function (window, undefined) {
    "use strict";

    var extend = Q.extend,
        getType = Q.type,

        toSet = Q.toSet,

        Num = Q.Num,
        Num_add = Num.add,
        Num_sub = Num.sub,
        Num_mul = Num.mul,
        Num_div = Num.div,

        SS = Q.SS;

    var START_DATE_FOR_EXCEL = new Date(1899, 11, 30),       //只读,1899/12/30
        MILLISECONDS_OF_DAY = 86400000,                      //1天表示的毫秒数

        hour12_suffixs = ["A/P", "AM/PM", "上午/下午"],      //12小时制后缀列表
        hour12_suffixs_len = hour12_suffixs.length,

        list_hour12_suffix = hour12_suffixs.map(function (suffix) {
            return [suffix, suffix.length, new RegExp("^" + suffix + "$", "i"), suffix.split("/")];
        }),
        map_hour12_suffix = hour12_suffixs.toMap(function (suffix) {
            return [suffix.charAt(0), true];
        }),

        RE_FORMAT_DATE = new RegExp("[ymdhsb]|a{3,}|" + hour12_suffixs.join("|") + "|e(?![+-])", "i"),
        RE_FORMAT_GENERAL = /^general$/i,                                       //正则:检测是否为常用格式,不区分大小写
        RE_FORMAT_COLOR = /^Color/i,                                            //正则:检测是否为颜色
        RE_FORMAT_SPECIALNUM = /^DBNum/i,                                       //正则:检测是否为特殊格式(中文大小写等)
        RE_FORMAT_SYMBOL = /^\$([^-]*)-([a-zA-Z0-9]+)$/,                        //正则:检测货币符号或区域性标识符
        RE_FORMAT_RULE = /^(>|>=|=|<|<=|<>)((-|0|[1-9])\d*(\.\d+)?)$/,          //正则:检测逻辑条件
        RE_ONEBYTE_CHAR = /[\x00-\xff]/,                                        //正则:检测单字节字符
        RE_QUOTE_STRING = /"\\""/g,                                             //正则:编码后的引号
        RE_QUOTE_CONTENT = /"[^"]+"/g,                                          //正则:引号内容
        RE_FORMAT_SCIENTIFICNUM = /^([0-9\.]+)E([+-])(\d+)$/,                   //正则:科学记数
        RE_FORMAT_STARTZERO = /^[0]+/,                                          //正则:开始处的0
        RE_FORMAT_TEXTFORMAT = /^[^#\.0\?bdeghmnsy%]+$/i,                       //正则:检测是否为文本格式

        map_date_char = toSet("ymdhsateYMDHSAT"),                               //日期字符
        map_error_number_char = toSet("bdeghmnsy[]"),                           //导致数字格式错误的字符
        map_number_char = toSet("#,.0?"),                                       //正确的数字格式字符
        map_named_color = toSet(["black", "red", "white", "blue", "cyan", "green", "yellow", "magenta"]),  //已命名颜色

        hiddenFormat = { isHide: true };                                        //隐藏格式

    //星期单位
    var list_week_unit = ["周", "週"];

    //月份名称列表
    var list_month_source = [
        ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    ];

    //星期名称列表
    var list_week_source = [
        ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
        ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
    ];

    //特殊字符串源数组
    var map_special_source = {
        "1": ["○", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "百", "千", "万", "亿", "兆"],
        "2": ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖", "拾", "佰", "仟", "万", "亿", "兆"],
        "3": ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "十", "百", "千", "万", "亿", "兆"],
        "4": ["○", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "百", "千", "萬", "億", "兆"],
        "5": ["零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖", "拾", "佰", "仟", "萬", "億", "兆"]
    };

    //日期与时间格式
    var DateTimeFormat = {
        LONG_DATE: "yyyy年M月d日",
        SHORT_DATE: "yyyy/M/d",
        LONG_TIME: "hh:mm:ss",
        SHORT_TIME: "h:mm:ss"
    };

    //将数字转化为日期对象
    function numberToDate(value) {
        var days = typeof value == "number" ? value : parseFloat(value), startDate = new Date(1899, 11, 30);

        if (Math.floor(days) == days) startDate.setDate(30 + days);
        else startDate.setMilliseconds(Math.round(days * MILLISECONDS_OF_DAY));

        return startDate;
    }

    //将日期对象转化为数字
    function dateToNumber(date) {
        if (typeof date == "number") return date;
        if (getType(date) != "date") date = Date.from(date);
        return (date - START_DATE_FOR_EXCEL) / MILLISECONDS_OF_DAY;
    }

    //去除引号内容(包括引号)
    //format:经过 CellFormat.parse 解析过后的格式
    //RE_QUOTE_STRING:解析过后的引号 eg:!" => "\""
    //RE_QUOTE_CONTENT:一般引号内容
    function dropQuoteOfParsedFormat(format) {
        return format.replace(RE_QUOTE_STRING, "").replace(RE_QUOTE_CONTENT, "");
    }

    //是否日期格式
    function isDateFormat(format) {
        return format != "" && RE_FORMAT_DATE.test(dropQuoteOfParsedFormat(format));
    }

    //是否为分数格式
    function isFractionFormat(format) {
        return format != "" && /\//.test(dropQuoteOfParsedFormat(format));
    }

    //是否为文本格式
    function isTextFormat(format) {
        return format != "" && RE_FORMAT_TEXTFORMAT.test(dropQuoteOfParsedFormat(format));
    }

    //转化文本(无单位 eg:2009 => 二〇〇九)
    function getSpecialText(value, specialSource) {
        if (typeof value != "string") value = value.toString();

        if (!specialSource) return value;

        var tmp = [];
        for (var i = 0, len = value.length; i < len; i++) {
            var c = value.charAt(i);
            tmp.push(specialSource[c] || c);
        }
        return tmp.join("");
    }

    //转化日期数字
    function getSpecialDate(value, len, source) {
        if (source) {
            if (value > 19) {
                var result = source[Math.floor(value / 10)] + source[10], n = value % 10;
                return n != 0 ? result + source[n] : result;
            }
            if (value > 10) return source[10] + source[value - 10];
            if (value == 10) return source[10];
            return len == 1 ? source[value] : source[0] + source[value];
        }

        var result = value.toString();

        return len != 1 && value < 10 ? "0" + result : result;
    }

    //转化整数(有单位 eg:2009 => 二千〇九)
    function getSpecialInt(value, specialSource) {
        if (typeof value != "string") value = value.toString();

        if (!specialSource) return value;

        var tmp = [], len = value.length, count = 0, skip = false;
        for (var i = len - 1; i >= 0; i -= 4) {
            //处理单位
            if (count > 0) {
                var unit = count % 3;
                if (unit == 0) unit = 3;
                tmp.unshift(specialSource[12 + unit]); //万、亿、兆
            }

            //以4位数一组进行处理
            var index = i - 4;
            if (index < -1) index = -1;

            for (var j = i; j > index; j--) {
                var c = value.charAt(j);
                if (c == "0") {
                    //若最后一位为0,则跳过 eg:2000
                    if (j == i) {
                        skip = true;
                        continue;
                    }

                    //若前面一位也是0,则跳过 eg:20[0]0 ->跳过,2[0]1 ->不跳过,继续向下处理
                    if (j - 1 > index && value.charAt(j - 1) == "0") continue;
                    if (!skip) tmp.unshift(specialSource[0]);
                } else {
                    skip = false;
                    if (i - j > 0) tmp.unshift(specialSource[9 + i - j]);  //十、百、千
                    tmp.unshift(specialSource[c] || c);     //倒序,所以放在单位(十、百、千)后面
                }
            }
            count++;
        }
        return tmp.join("");
    }

    //===============================日期格式解析=============================
    //构造器:日期格式对象
    function DateFormat(ops) {
        if (ops) extend(this, ops);

        this.isDate = true;

        this.init();
    };

    DateFormat.prototype = {
        //初始化
        init: function () {
            var that = this, source = that.source;

            if (!source) {
                source = "";
                if (!that.color && !that.isGeneral && !that.result) that.isHide = true;
            } else if (RE_FORMAT_GENERAL.test(source)) {
                that.isGeneral = true;
            } /*else {
                switch (source) {
                    case "m/d/yy": source = DateTimeFormat.SHORT_DATE; break;
                    case "m/d/yy h:mm": source = "yyyy/m/d h:mm"; break;
                }
            }*/

            switch (this.LCID) {
                case "f800": source = DateTimeFormat.LONG_DATE; break;
                case "f400": source = DateTimeFormat.SHORT_TIME; break;
            }

            that.source = source;

            if (that.isHide || that.isGeneral) return;
            this.parse(source);
        },
        //格式解析
        parse: function (source) {
            var list = [], fs = [], indexs = {}, len = source.length, quote = 0, index = 0, str = "", isInBracket = false, isNotFirst = false, isHour12 = false;

            for (var i = 0; i < len; i++) {
                var c = source.charAt(i);

                //处理引号内容
                if (c == '"') {
                    if (i == 0 || source.charAt(i - 1) != "\\") {
                        quote++;
                        if (quote == 1) index = i + 1;
                        else {
                            if (i > index) {
                                var s = source.slice(index, i);
                                list.push(s == '\\"' ? '"' : s);
                            }
                            quote = 0;
                        }
                        continue;
                    }
                }

                //跳过引号内容
                if (quote > 0) continue;

                //处理[],eg:[h],[m],[s]
                if (c == "[") {
                    index = i + 1;
                    isInBracket = true;
                    continue;
                }

                if (c == "]") {
                    var key = source.slice(index, i).toLowerCase(), unit = key.charAt(0);
                    list.push({ key: key, fn: "total", args: [unit, isNotFirst] });

                    isInBracket = false;
                    continue;
                }

                if (isInBracket) continue;

                //处理12小时制名称 eg:AM/PM
                if (map_hour12_suffix[c]) {
                    for (var k = 0; k < hour12_suffixs_len; k++) {
                        var sl = list_hour12_suffix[k], _str = source.substr(i, sl[1]);
                        if (sl[2].test(_str)) {
                            if (!isHour12) isHour12 = true
                            list.push({ key: sl[0], fn: "getHour12Name", args: [sl[3]] });
                            i += sl[1] - 1;
                            break;
                        }
                    }
                    continue;
                }

                //不参与计算的字符 eg:-/: etc
                if (!map_date_char[c]) {
                    list.push(c);
                    continue;
                }

                isNotFirst = true;

                //若当前字符与相邻的下一个字符相同,则将当前字符串追加到临时字符串并跳出循环
                if (i != len - 1 && c == source.charAt(i + 1)) {
                    str += c;
                    continue;
                }

                //若当前字符与相邻的上一个字符相同,则将当前字符串追加到临时字符串
                //否则直接赋值给临时字符串
                if (i > 0 && c == source.charAt(i - 1)) str += c;
                else str = c;

                var firstChar = str.toLowerCase().charAt(0);

                //处理e,使用4位数显示年份,不参与月份(M)、分钟(m)识别
                if (firstChar == "e") {
                    list.push({ key: "e", fn: "getYear", args: [4] });
                    str = "";
                    continue;
                }

                if (firstChar == "a" && str.length < 3) {
                    //将a或aa作为字符串处理
                    list.push(str);
                    str = "";
                    continue;
                }

                indexs[list.length] = fs.length;
                fs.push(str);
                str = "";
            }  //end for

            //月份(M)、分钟(m)识别及其它处理(12小时制等)
            var isMonth = true,    //是否月份
                skip = false,      //是否跳过s的影响
                sCount = 0,        //s出现的次数
                fList = [];

            for (var i = 0, fsLen = fs.length; i < fsLen; i++) {
                var str = fs[i], _len = str.length, c = str.charAt(0).toLowerCase(), fn, args;

                if (c == "h") {
                    isMonth = false;
                    skip = true;
                } else if (c == 's') {
                    if (skip) skip = false;
                    else if (sCount == 0) isMonth = false;  //如果不止一个s,则该s不再影响后面的m

                    sCount++;
                } else if (c == 'm') {
                    if (isMonth) {
                        //如果m后面紧跟s,则m为分钟
                        if (i + 1 < fsLen && fs[i + 1].charAt(0) == 's') {
                            isMonth = false;
                            skip = true; //s作用于前一个m,后面的m不再受此s的影响
                        }
                    }
                }

                //封装处理
                switch (c) {
                    case "y": fn = "getYear", args = [_len]; break;
                    case "m":
                        if (_len > 2) {
                            //英文月份
                            //mmm:使用英文缩写来显示月份(Jan～Dec)
                            //mmmm:使用英文全称来显示月份(January～December)
                            //mmmmm:显示月份的英文首字母(J～D)
                            fn = "getMonthName", args = [0];
                            if (_len == 3) args.push(3);
                            else if (_len == 5) args.push(1);
                        } else {
                            //月份与分钟
                            args = [_len];
                            if (isMonth) {
                                fn = "getMonth";
                            } else {
                                fn = "getMinutes";
                                isMonth = true;
                            }
                        }
                        break;
                    case "d":
                        if (_len < 3) {
                            //日期
                            fn = "getDate", args = [_len];
                        } else {
                            //英文星期
                            fn = "getWeek";     //fn = "getWeek", args = _len == 3 ? [0, 0, 3] : [0];
                            if (this.LCID == "804") args = _len == 3 ? [1, -1, undefined, list_week_unit[0]] : [1];
                            else if (this.LCID == "404") args = _len == 3 ? [1, -1, undefined, list_week_unit[1]] : [1];
                            else args = _len == 3 ? [0, 0, 3] : [0];
                        }
                        break;
                    case "a":
                        //中文星期
                        fn = "getWeek";         //fn = "getWeek", args = _len == 3 ? [1, -1] : [1];
                        if (this.LCID == "409") args = _len == 3 ? [0, 0, 3] : [0];
                        else if (this.LCID == "404") args = _len == 3 ? [1, -1, undefined, list_week_unit[1]] : [1];
                        else args = _len == 3 ? [1, -1, undefined, list_week_unit[0]] : [1];

                        break;
                    case "h":
                        //小时
                        fn = "getHours", args = [_len, isHour12];
                        break;
                    case "s":
                        //秒
                        fn = "getSeconds", args = [_len];
                        break;
                }

                fList.push({ key: str, fn: fn, args: args });
            }  //end for

            this.source = source;
            var output = [];
            for (var i = 0, len = list.length; i < len; i++) {
                if (indexs[i] != undefined) output.push(fList[indexs[i]]);
                output.push(list[i]);
            }
            if (indexs[list.length] != undefined) output.push(fList[indexs[list.length]]);

            this.list = output;
        },
        //输出格式后的内容
        format: function (n) {
            var num = n;
            if (isNaN(num)) num = dateToNumber(num);
            if (isNaN(num)) return n;

            //Excel 日期里1900年2月有29天
            //60 => 1900/2/29
            var date_ms = num, is_1900_2_29 = false;
            if (num > 0 && num < 61) {
                if (num == 60) is_1900_2_29 = true;
                else num++;
            }

            //负数(值为负数时,日期、时间不解析)
            //if (num < 0) return -1; //{ fill: "#" };

            var date = numberToDate(num), specialNum = this.specialNum, output = [], list = this.list || [], util = DateFormat.util, specialSource = specialNum ? map_special_source[specialNum] : null;

            util.source = specialNum ? map_special_source[specialNum] : null;
            util.date = date;
            util.num = date_ms;
            util.is_1900_2_29 = is_1900_2_29;

            for (var i = 0, len = list.length; i < len; i++) {
                var f = list[i];
                if (typeof f == "string") output.push(f);
                else output.push(util[f.fn].apply(util, f.args));
            }

            return output.join("");
        }
    };

    //工具方法
    DateFormat.util = {
        //获取年份
        getYear: function (len) {
            var year = this.date.getFullYear().toString();
            return getSpecialText(year.slice(-len), this.source);
        },
        //获取月份
        getMonth: function (len) {
            return getSpecialDate(this.date.getMonth() + 1, len, this.source);
        },
        //获取月份名称
        getMonthName: function (index, len) {
            var source = list_month_source[index], month = this.date.getMonth(), name = source[month]; //source 索引从0开始,所以月份不需加1
            return len ? name.slice(0, len) : name;
        },
        //获取日期
        getDate: function (len) {
            var day = this.date.getDate();
            if (this.is_1900_2_29) day++;

            return getSpecialDate(day, len, this.source);
        },
        //获取星期
        getWeek: function (index, first, last, weekUnit) {
            var source = list_week_source[index], week = this.date.getDay(), name = source[week];
            if (first == undefined && last == undefined) return name;
            if (weekUnit == undefined) weekUnit = "";

            return last == undefined ? weekUnit + name.slice(first) : weekUnit + name.slice(first, last);
        },
        //获取小时
        getHours: function (len, isHour12) {
            var hour = this.date.getHours();
            if (isHour12) {
                if (hour > 12) hour -= 12;
                if (hour == 0) hour = 12;
            }
            //console.log(this.date)
            return getSpecialDate(hour, len, this.source);
        },
        //获取12小时制名称
        getHour12Name: function (names) {
            return names[this.date.getHours() > 12 ? 1 : 0];
        },
        //获取分钟
        getMinutes: function (len) {
            return getSpecialDate(this.date.getMinutes(), len, this.source);
        },
        //获取秒钟
        getSeconds: function (len) {
            return getSpecialDate(this.date.getSeconds(), len, this.source);
        },
        //显示超出进制的时间 eg:[d]、[h]、[m]、[s]
        total: function (unit, isNotFirst) {
            var date = numberToDate(this.num), result;

            if (isNotFirst) {
                result = 12 + date.getSeconds() * 24;
                switch (unit) {
                    case "m": result *= 60; break;
                    case "s": result *= 3600; break;
                }
            } else {
                result = Math.round(Date.total(unit, date - START_DATE_FOR_EXCEL));
            }

            return getSpecialInt(result, this.source);
        }
    };

    //==============================数字格式解析==========================
    //author:zhangxin
    //数字解析
    function NumberFormat(ops) {
        if (ops) extend(this, ops);

        this.init();
    }

    //数字解析构造器
    NumberFormat.prototype = {
        //初始化
        init: function () {
            var that = this, source = that.source;
            if (!source) {
                source = that.source = "";
                if (!that.color && !that.isGeneral && !that.result) that.isHide = true;
            } else if (RE_FORMAT_GENERAL.test(source)) that.isGeneral = true;

            if (that.isHide || that.isGeneral) return;
            that.parse(source);
        },
        //解析数字
        parse: function (source) {
            var that = this, arr = [], quote = 0,
                dotLength = 0,          //小数长度
                powerLength = 0,        //次幂长度(科学记数)
                integerLength = 0,      //整数长度
                isFirstFormat = true;   //是否为第一个参与计算的字符

            that.isFractionFormat = isFractionFormat(source);

            //分数(把分数格式获取并替换为F)
            if (that.isFractionFormat) {
                var vs = dropQuoteOfParsedFormat(source).split('/'), leftSource = vs[0] || "", rightSouce = vs[1] || "", fraction_source = "";
                that.radix = 0;                     //分母位数
                that.denominator = "";      //分母(数字)

                //获取分子格式
                for (var len = leftSource.length, l = len - 1; l >= 0; l--) {
                    var c = leftSource.charAt(l) || "";
                    if (c == ' ' || c == '"') {
                        l = -1;
                        continue;
                    } else fraction_source = c + fraction_source;
                }

                //获取分母格式
                for (var r = 0, len = rightSouce.length; r < len; r++) {
                    var c = rightSouce.charAt(r) || " ";
                    if (r == 0) fraction_source += "/";

                    if (c == ' ' || c == '"') continue;
                    else {
                        that.radix++;                           //分母位数
                        if (!isNaN(c)) that.denominator += c;   //分母为纯数字
                        fraction_source += c;
                        continue;
                    }
                }
                source = source.replace(fraction_source, 'F');
                that.fractionFormat = fraction_source;          //分数格式

                //解析起作用的整数格式
                var integer_format = source.split("F")[0], isInteger = true, _newFormat = "", quote = 0, index = 0, countBlank = 0;
                for (var len = integer_format.length, i = len - 1; i >= 0; i--) {
                    var c = integer_format.charAt(i) || "";
                    if ((c == "#" || c == "?" || c == "0") && isInteger) {
                        if (i - 1 >= 0 && integer_format.charAt(i - 1) == c) {
                            if (c == "?" || c == "0") _newFormat = c + _newFormat;
                            continue;
                        } else {
                            _newFormat = c + _newFormat;
                            isInteger = false;
                            continue;
                        }
                    }

                    //?
                    if (c == '?' && !isInteger) {
                        _newFormat = " " + _newFormat;
                        continue;
                    }

                    //处理引号内容
                    if (c == '"') {
                        quote++;
                        if (quote == 1) index = i + 1;
                        else {
                            if (i > index) {
                                var s = source.slice(index, i);
                                _newFormat = c + _newFormat;
                            }
                            quote = 0;
                        }
                        continue;
                    }

                    //#
                    if (c != "#") {
                        _newFormat = c + _newFormat;
                        continue;
                    }
                }
                source = source.replace(integer_format, _newFormat); //xlog(integer_format+"||"+_newFormat)
            }

            for (var i = 0, len = source.length; i < len; i++) {
                var c = source.charAt(i);

                //处理引号内容
                if (c == '"') {
                    if (i == 0 || source.charAt(i - 1) != "\\") {
                        quote++;
                        if (quote == 1) index = i + 1;
                        else if (quote == 2) {
                            var s = source.slice(index, i);
                            arr.push({ content: (s == '\\"' ? '"' : s) });  //直接输出的字符
                            quote = 0;
                        }
                        continue;
                    }
                }

                //跳过引号内容
                if (quote > 0) continue;

                //跳过*
                if (c == "*") continue;

                //是否为General
                if (c == "G") {
                    that.hasGeneral = true;
                    arr.push(c);
                    continue;
                }

                //千分位
                if (c == ",") {
                    that.hasComma = true;
                    arr.push(c);
                    continue;
                }

                //百分率
                if (c == "%") {
                    that.hasPercent = true;
                    arr.push({ content: c });
                    continue;
                }

                //小数
                if (c == ".") {
                    that.hasDot = true;
                    arr.push(c);
                    continue;
                }

                //@(文本占位符)
                if (c == "@") {
                    arr.push(c);
                    continue;
                }

                //分数
                if (that.isFractionFormat && (c == "F")) {
                    arr.push(c);
                    continue;
                }

                //科学记数
                if (c == "E") {
                    arr.push({ content: "" }); //(c);
                    that.isScientificNum = true;

                    if (i + 1 < len) {
                        var nextChar = source.charAt(i + 1);
                        if (nextChar == "+" || nextChar == "-") {
                            arr.push(nextChar);
                            i++;
                        }
                    } //end if
                    continue;
                } //end if

                //参与计算的字符
                if (map_number_char[c]) {
                    if (that.isScientificNum) powerLength++;    //科学记数的次幂长度
                    else if (that.hasDot) dotLength++;          //小数的长度
                    else integerLength++;      //整数部分的长度

                    if (isFirstFormat) {    //第一个参与计算的字符
                        arr.push({ isFirstFormat: true });
                        isFirstFormat = false;
                    }
                    arr.push(c);
                    continue;
                } else arr.push({ content: c }); //不参与计算的字符

            } //end for
            that.source = source;
            that.dotLength = dotLength;                 //小数位数
            that.powerLength = powerLength;             //次幂位数(科学记数)
            that.integerLength = integerLength;         //整数长度(科学记数)
            that.list = arr;
        },
        //按格式输出
        format: function (num) {
            var that = this,
                list = that.list || [],                     //格式
                hasComma = that.hasComma,                   //是否存在千分位
                hasDot = that.hasDot,                       //是否存在小数
                isScientificNum = that.isScientificNum,     //是否为科学记数
                dotLength = that.dotLength,                 //保留小数长度
                powerLength = that.powerLength,             //次幂长度(科学记数)
                integerLength = that.integerLength,         //整数长度
                specialNum = that.specialNum,               //特殊字符
                isGeneral = that.isGeneral,                 //是否为标准
                hasGeneral = that.hasGeneral,               //是否存在标准
                isMinus = false,                            //是否为负数
                util = NumberFormat.util;

            //参数处理
            if (isNaN(num)) return num;             //非数字

            //分数(分母位数长度大于5不解析，因为轮循算法找最近似值太慢)
            if (that.isFractionFormat && that.radix > 5) return num;

            if (num < 0) {
                isMinus = true;
                num = -num;     //转为正数
            }
            if (that.isMinus) isMinus = false;

            //百分率
            if (that.hasPercent) num = Num_mul(num, 100);

            var value = num + "";
            if (isScientificNum) value = util.parseSicenceNum(num, integerLength, dotLength, powerLength, hasDot);       //科学记数

            var vs = value.split('.'),
                wholeNumber = vs[0] || "",
                dotNumber = vs[1] || "";          //整数部分,小数部分

            //当数字的小数位数小于格式保留的位数
            if (dotLength > dotNumber.length && !isScientificNum) {
                dotNumber += "0".repeat((dotLength - dotNumber.length) || 0);
            }

            //中文大小写
            if (specialNum > 0 && !isScientificNum) {
                if (that.LCID == "404") specialNum = (parseInt(specialNum) + 3).toString();              //地区:台湾

                if (hasGeneral || isGeneral) wholeNumber = getSpecialInt(wholeNumber, map_special_source[specialNum]);      //标准的格式时,中文大小写带单位
                else wholeNumber = getSpecialText(wholeNumber, map_special_source[specialNum]);                             //标准的格式时,中文大小写不带单位
                if (dotNumber.length > 0) dotNumber = getSpecialText(dotNumber, map_special_source[specialNum]);
            }

            //小数
            if (dotLength < dotNumber.length && !isScientificNum && !hasGeneral && !isGeneral) {
                /*var roundNum = dotNumber.charAt(dotLength);
                //不保留小数
                if (dotLength == 0) {
                    if (roundNum > 4) wholeNumber = (parseInt(wholeNumber) + 1).toString();
                    dotNumber = "";
                } else { //小数部分
                    if (roundNum > 4) {     //四舍五入
                        dotNumber = (parseInt(dotNumber.substring(0, dotLength)) + 1).toString();
                        if (dotNumber.length > dotLength) {
                            wholeNumber = (parseInt(wholeNumber) + 1).toString();
                            dotNumber = 0;
                        }
                    } else dotNumber = dotNumber.substring(0, dotLength);
                }*/

                var roundNum = parseInt(dotNumber.charAt(dotLength));
                //不保留小数
                if (dotLength == 0) {
                    if (roundNum > 4) wholeNumber = (parseInt(wholeNumber) + 1).toString();
                    dotNumber = "";
                } else { //小数部分
                    var retentionDecimal = dotNumber.substring(0, dotLength), _index = 0, _len = 0;
                    if (roundNum > 4) {
                        retentionDecimal = (parseInt(retentionDecimal) + 1).toString(), _len = retentionDecimal.length;
                        if (_len > dotLength) {
                            wholeNumber = (parseInt(wholeNumber) + 1).toString();
                            _index = 1;
                        } else retentionDecimal = "0".repeat(dotLength - _len) + retentionDecimal;
                    }
                    dotNumber = _index > 0 ? retentionDecimal.substring(1) : retentionDecimal;
                }
            }

            if (hasComma && !isScientificNum) wholeNumber = util.addComma(wholeNumber);                 //千分位处理

            var newNum = dotNumber.length == 0 ? wholeNumber : (wholeNumber + "." + dotNumber),
                numLen = newNum.length,
                isSkipZero = false,          //是否跳过0
                isSkipQuestion = false,      //是否跳过?(用于分数格式整数输出)
                str = "";

            //按格式输出数字
            for (var len = list.length, i = len - 1; i >= 0; i--) {
                var c = list[i], strC = "";

                if (typeof c == "object") {
                    if (c.isFirstFormat == undefined) strC = c.content || "";                 //直接输出的字符串
                    else {          //输出剩下的数字
                        if (that.isFractionFormat) continue;
                        strC = newNum.substring(0, numLen) || "";
                    }
                } else if (c == "@") {
                    //出现一次@重复一次值
                    strC = newNum;
                } else {
                    numLen--;
                    var cutChar = newNum.charAt(numLen);
                    if (c == "." || c == ",") strC = cutChar;
                    else if (c == "#") {                                                            //#
                        if (that.isFractionFormat) strC = newNum == "0" ? "" : newNum;
                        else strC = numLen >= 0 ? cutChar : "";
                    }
                    else if (c == "0" || c == "?") {                                                //0&?
                        if (that.isFractionFormat) {
                            if (c == "0" && !isSkipZero) {
                                strC = integerLength > newNum.length ? "0".repeat(integerLength - newNum.length) + newNum : newNum;
                                isSkipZero = true;
                            }

                            if (c == "?" && !isSkipQuestion) {
                                strC = integerLength > newNum.length ? (" ").repeat(integerLength - newNum.length) + newNum : newNum;
                                isSkipQuestion = true;
                            }
                        } else strC = numLen >= 0 ? cutChar : c;
                    } else if (c == "E") strC = cutChar;                                            //科学记数
                    else if (c == "G" && hasGeneral) strC = newNum;                                 //存在General
                    else if (c == "F" && that.isFractionFormat) {
                        if (that.denominator == "") strC = integerLength > 1 ? "  " : " " + util.decimalToFraction(value, (integerLength < 1), that.radix);     //分数
                        else {                                                                                          //分母为指定数字
                            var denominator = parseInt(that.denominator), numerator = util.numToFraction(value, denominator);
                            strC = numerator == 0 ? "   " : numerator + "/" + that.denominator;
                        }

                    } else strC = cutChar;
                }

                str = strC + str;
            }

            if (list.length == 0) str = newNum;

            return (isMinus ? "-" : "") + str;
        }
    };

    //工具方法
    NumberFormat.util = {
        //给整数部分添加千分位
        addComma: function (num) {
            num = num + "";
            var vs = num.split("."),
                wholeNumber = vs[0] || "",
                dotPart = vs[1] || "";

            if (wholeNumber.length > 3) {
                var mod = num.length % 3;
                wholeNumber = mod == 0 ? "" : num.substring(0, mod);

                for (var i = 0, len = Math.floor(num.length / 3) ; i < len ; i++) {
                    if ((mod == 0) && (i == 0)) wholeNumber += num.substring(mod + 3 * i, mod + 3 * i + 3);
                    else wholeNumber += ',' + num.substring(mod + 3 * i, mod + 3 * i + 3);
                }
            }

            return wholeNumber + dotPart;
        },
        //判断是否为科学记数
        isSicenceNum: function (format) {
            return RE_FORMAT_SCIENTIFICNUM.test(format);
        },
        //把数字转化为科学记数(value-数字,integralLength-整数长度,dotLength-小数位数,powerLength-次幂位数)
        parseSicenceNum: function (value, integralLength, dotLength, powerLength, isDot) {
            if (RE_FORMAT_SCIENTIFICNUM.test(value)) return value;  //是否为科学记数
            var num = value, isMinus = false, integralLength = integralLength > 3 ? 3 : integralLength || 1, dotLength = dotLength >= 0 ? dotLength : 2, powerLength = powerLength || 2;
            //是否为负数
            if (num < 0) {
                isMinus = true;
                num = -num;
            }
            var str = num.toString(), index = 0, fix = 0, count = 0, unit = "+", dotIndex = str.indexOf("."), count = 0, num_length = str.length;

            //存在小数
            if (dotIndex >= 0) {
                if (dotIndex == 1) {    //整数部分只有一位数(如:1.1210、0.00153)
                    var firstChar = str.charAt(0);
                    if (firstChar == "0") {
                        var arr_num = RE_FORMAT_STARTZERO.exec(str.substring(dotIndex + 1));
                        fix = arr_num != null ? (arr_num[0].length + 1) : 1;
                        index = num_length - dotIndex - 1;
                        count = index - fix;
                        unit = "-";
                    } else if (parseInt(firstChar) > 0) {
                        fix = index = count = 0;
                    }
                } else if (dotIndex > 1) {      //整数部分有多位数字(如:1286.12)
                    index = num_length - dotIndex - 1;
                    count = num_length - 2;
                    fix = count - index;
                }
            } else {                     //整数
                index = 0;
                if (integralLength >= num_length) fix = 0;
                else fix = num_length - integralLength;
            }

            count = count == 0 ? fix : count;
            if (integralLength > 0 && count >= integralLength) count = count - integralLength + 1;
            else count = 0;

            num = Num_div(Num_mul(num, Math.pow(10, index)), Math.pow(10, count)).toString();

            //保留小数的位数
            if (dotLength > 0) {
                var _index = num.indexOf("."),                  //小数点的位置索引
                        _numIndex = _index > 0 ? _index : 0,
                        _dotNumber = num.length - _numIndex - 1;  //数值的小数部分长度

                //小数部分位数大于需要保留长度
                if (_dotNumber > dotLength) {
                    var _num = 0;
                    if (dotLength + 1 <= _dotNumber) _num = num.slice(_numIndex + 1, (dotLength + 2 + _numIndex));
                    num = num.slice(0, _numIndex + 1) + Math.round(Num_div(parseInt(_num), 10)).toString();
                } else if (_dotNumber == dotLength) {                       //小数部分位数等于需要保留长度
                    num = num;
                } else if (_dotNumber < dotLength) {                        //小数部分位数小于需要保留长度
                    if (_numIndex == 0) num += ".";
                    num += "0".repeat(dotLength - _dotNumber);
                }
            } else if (dotLength == 0) num = num.slice(0, 1);

            //保留次幂的位数
            if (fix != 0) {
                if (unit == "-") fix = fix + integralLength - 1;
                else fix = fix - integralLength + 1;
            }

            var fix_len = fix.toString().length;
            if (powerLength > fix_len) {
                fix = "0".repeat(powerLength - fix_len) + fix;
            }
            if (isMinus) num = -num;

            return num + ((isDot && dotLength == 0) ? '.' : "") + "E" + unit + fix;
        },
        //小数转化为分数
        decimalToFraction: function (value, hasInteger, radix) {
            var minError = 1, denominator = 0;                  //最小误差、分母值
            if (value < 0) value = Math.abs(value);

            var integer = Math.floor(value), decimal = Num_sub(value, integer), maxNumerator = Math.ceil(Num_mul(decimal, Math.pow(10, radix))), //整数值、小数值、分子最大值
                currentDenominator, currentError;               //当前分母、当前误差值
            if (decimal == 0) return "";              //整数

            for (var i = 1; i < maxNumerator; i++) {
                currentDenominator = Math.round(i / decimal);
                currentError = Math.abs(i / currentDenominator - decimal);
                if (minError > currentError && currentDenominator < Math.pow(10, radix)) {
                    minError = currentError;
                    denominator = currentDenominator;
                }
            }

            return Math.round(denominator * decimal) + (hasInteger ? denominator * integer : 0) + "/" + denominator;
        },
        //小数转化为分数(指定分母值)
        numToFraction: function (value, denominator) {
            if (value < 0) value = Math.abs(value);
            var decimal = Num_sub(value, Math.floor(value));

            return Math.round(Num_mul(decimal, denominator));
        }
    };

    //============================文本格式解析======================
    function TextFormat(ops) {
        if (ops) extend(this, ops);

        this.init();    //初始化
    }

    TextFormat.prototype = {
        //初始化
        init: function () {
            var that = this, source = that.source;
            if (!source) that.isHide = true;
            else if (RE_FORMAT_GENERAL.test(source)) that.isGeneral = true;

            if (that.isHide || that.isGeneral) return;
            that.parse(source);
        },
        //解析
        parse: function (source) {
            var that = this, arr = [], count = 0, quotationMarks = 0;
            for (var i = 0, len = source.length; i < len; i++) {
                var c = source.charAt(i);
                //下划线'_'
                if (c == "_") {
                    arr.push('" "');
                    i++;
                    continue;
                }

                //文本占位符'@'
                if (c == '@') {
                    if (i + 1 < len && source.charAt(i + 1) == c) count++;
                    else {
                        arr.push(c);
                        count = 0;
                    }
                    continue;
                }

                //空格' '
                if (c == " ") {
                    arr.push(c);
                    continue;
                }
            }

            that.list = arr;
        },
        //按格式输出
        format: function (value) {
            var that = this, list = that.list, len = list.length;
            if (len == 0) return value;

            var str = "";
            for (var i = 0; i < len; i++) {
                if (list[i] == '@') str += value;
                else str += list[i];
            }
            return str;
        }
    };

    //==============================条件规则========================
    //条件规则
    function Rule(key, value) {
        this.key = key;
        this.value = parseFloat(value);
    }

    //条件规则构造器
    Rule.prototype = {
        //规则匹配
        match: function (value) {
            switch (this.key) {
                case ">": return value > this.value;
                case "<": return value < this.value;
                case "=": return value == this.value;
                case ">=": return value >= this.value;
                case "<=": return value <= this.value;
                case "<>": return value != this.value;
            }
            return false;
        }
    };

    //==============================单元格格式解析========================
    //构造器:单元格格式对象
    function CellFormat(source) {
        this.source = source;   //单元格格式 eg:"￥"#,##0.00_);[Red]\("￥"#,##0.00\)
        this.list = [];
        this.rules = [];

        this.parse(source);
    }

    CellFormat.prototype = {
        //格式解析
        parse: function (source) {
            var that = this,
                len = source.length,
                lastTwo = len - 2,
                isGeneral = false;

            //若以";@"结尾,则把";@"去掉
            if (lastTwo >= 0 && source.slice(lastTwo) == ";@") source = source.slice(0, lastTwo);

            //空格式或常用格式(General)
            if (source == "" || RE_FORMAT_GENERAL.test(source)) {
                that.isGeneral = true;
                return;
            }

            var list = that.list,
                rules = that.rules,
                tmp = [],
                quote = 0,
                index = 0,
                isNotInBracket = true,     //是否非中括号内容
                isSegment = false,
                bf = {};

            for (var i = 0; i < len; i++) {
                var c = source.charAt(i),         //当前字符
                    isNotInQuote = quote == 0,    //是否非引号内容
                    isPlaceChar = c == "_";       //是否是占位符"_"

                //!:强制显示下一个文本字符
                //\:显示后面的字符(和""用途相同)
                //_:留出与下一个字符等宽的空格
                if (c == "!" || c == "\\" || isPlaceChar) {
                    if (isNotInQuote && isNotInBracket) {
                        if (i > index) tmp.push(source.slice(index, i));

                        var nextChar = source.charAt(++i);
                        if (isPlaceChar) tmp.push(RE_ONEBYTE_CHAR.test(nextChar) ? '" "' : '"  "');   //空格
                        else tmp.push(nextChar == '"' ? '"\\""' : '"' + nextChar + '"');              //避免嵌套引号

                        index = i + 1;
                    }
                } else if (c == "[") {
                    if (isNotInQuote) {
                        if (i > index) tmp.push(source.slice(index, i));

                        index = i + 1;
                        isNotInBracket = false;
                    }
                } else if (c == '"') {
                    if (isNotInBracket) {
                        quote++;

                        if (quote == 2) quote = 0;
                    }
                } else {
                    //非双引号内容
                    if (isNotInQuote) {
                        if (isNotInBracket) {
                            switch (c) {
                                case ";":
                                    isSegment = true;
                                    break;
                                case "g": case "G":
                                    if (i > index) tmp.push(source.slice(index, i));
                                    if (RE_FORMAT_GENERAL.test(source.slice(i, i + 7))) {
                                        tmp.push("G");
                                        i += 6;
                                        index = i + 1;
                                    } else {
                                        tmp.push('"' + c + '"');
                                        index = i + 1;
                                    }
                                    isGeneral = true;
                                    break;
                                case "]":
                                    if (i > index) tmp.push(source.slice(index, i + 1));
                                    index = i + 1;

                                    //若没有"[",则将"]"转为引号内容
                                    tmp.push('"]"');
                                    break;
                            }
                        } else if (c == "]") {
                            //处理中括号内容
                            var value = source.slice(index, i).trim();
                            if (value != "") {
                                var _value = value.toLowerCase(), _str = _value.slice(0, 5);
                                if (map_named_color[_value]) bf.color = _value;                     //颜色(已命名颜色,不区分大小写) eg:red
                                else if (_str == "color") bf.color = value.slice(6);            //颜色(自定义颜色) eg:Color#00ffcc
                                else if (_str == "dbnum") bf.specialNum = value.slice(5);       //特殊格式(中文大小写等)
                                else if (RE_FORMAT_SYMBOL.test(value)) {
                                    //货币符号或区域性标识符
                                    var symbol = RegExp.$1, lcid = RegExp.$2;
                                    if (symbol) tmp.push('"' + symbol + '"');
                                    else if (lcid) bf.LCID = lcid.toLowerCase();
                                } else {
                                    //逻辑条件
                                    var match = RE_FORMAT_RULE.exec(value);
                                    if (match) rules.push(new Rule(match[1], match[2]));
                                    else if (_value == "h" || _value == "m" || _value == "s") tmp.push("[" + value + "]");
                                }
                            }  //end if (value != "")

                            index = i + 1;
                            isNotInBracket = true;
                        }  //end if (c == "]")
                    }  //end if (isNotInQuote)
                }  //end else

                var isLast = i >= len - 1;

                if (isLast || isSegment) {
                    if (isLast || isSegment) tmp.push(source.slice(index, isSegment ? i : len));

                    var format = that._format = tmp.join("");                //_format在设置格式模块中起检测是否为日期格式(负值不解析为时间格式)
                    if (format == "G") {
                        bf.isGeneral = true;
                        bf.source = "";
                    } else bf.source = format;

                    if (isDateFormat(format)) list.push(new DateFormat(bf));                                            //日期格式
                    else if (isTextFormat(format) && (!bf.isGeneral) && (!isGeneral)) list.push(new TextFormat(bf));    //文本格式
                    else list.push(new NumberFormat(bf));                                                               //数字格式

                    if (isLast && isSegment) list.push(hiddenFormat);

                    bf = {}, tmp = [];

                    if (isSegment) {
                        isSegment = false;
                        index = i + 1;
                    }
                }
            }
        },
        //获取匹配的格式
        getFormat: function (value) {
            var list = this.list, rules = this.rules, len = list.length, numOfRules = rules.length;
            if (len <= 0) return;

            //是否是数字
            var isNum = typeof value == "number";

            //当value为纯字符串时,返回value
            if (!isNum && len > 0) return list[0];

            //处理自定义规则
            if (numOfRules > 0) {
                var index = -1;
                if (isNum) {
                    for (var i = 0; i < numOfRules; i++) {
                        if (rules[i].match(value)) {
                            index = i;
                            break;
                        }
                    }

                    //条件匹配失败
                    if (index == -1) {
                        //规则长度大于2,则匹配第3个规则
                        if (len > 2) return list[2];

                        //规则长度小于或等于2,若规则与条件长度相同,则用#填充
                        if (len == numOfRules) return { result: -1 }; //{ result: "######" };

                        //匹配其它1个规则
                        return list[1];
                    }

                    return list[index];
                }

                //文本,若存在第4个规则,则匹配第4个规则,否则无论是否满足条件,皆匹配第1个规则
                return len == 4 ? list[3] : list[0];
            }

            //默认规则
            if (len == 1) return list[0];

            if (isNum) {
                if (value > 0) return list[0];  //正数(匹配第1个规则)
                if (value < 0) {                //负数(匹配第2个规则)
                    list[1].isMinus = true;
                    return list[1];
                }
                return list[len > 2 ? 2 : 0];      //0(匹配第3或第1个规则)
            }

            //若有4个规则,则文本匹配第4个规则,否则隐藏
            if (len == 4) return list[3];

            //隐藏文本
            return hiddenFormat;
        },
        //格式化输出
        //mode:为1时仅返回解析后的文本;为2时返回文本和颜色对象 {text,color};否则返回解析后的html(如果存在color)字符串
        format: function (value, mode) {
            if (this.isHide) return "";

            var instance = this.getFormat(value);
            if (!instance) return value.toString();

            if (instance.isHide) return "";
            if (instance.result != undefined) return instance.result;

            var color = instance.color, result = instance.format(value);

            switch (mode) {
                case 1: return result;
                case 2: return { text: result, color: color };
            }

            return color && typeof result == "string" ? '<span class="cell-format" style="color:' + color + ';">' + result + '</span>' : result || "";
        },
        //是否为日期格式
        isDate: function (value) {
            var instance = this.getFormat(value);
            return instance && instance.isDate;
        }
    };

    var map_cell_format = {};

    //获取单元格格式解析器
    function getCellFormatter(format) {
        if (!format) return undefined;

        var parser = map_cell_format[format];
        if (parser) return parser;

        parser = map_cell_format[format] = new CellFormat(format);

        return parser;
    }

    //------------------------ export ------------------------

    extend(SS, {
        numberToDate: numberToDate,
        dateToNumber: dateToNumber,

        CellFormatter: CellFormat,
        getCellFormatter: getCellFormatter
    });

})(window);

﻿/*
* cell-format.js
* author:devin87@qq.com
* update: 2016/01/15 11:05
*/
(function (window, undefined) {
    "use strict";


})(window);

﻿/*
* cell-formula.parse.js
* author:devin87@qq.com
* update: 2016/01/15 11:05
*/
(function (window, undefined) {
    "use strict";


})(window);

﻿/*
* cell-formula.js
* author:devin87@qq.com
* update: 2016/01/15 11:21
*/
(function (window, undefined) {
    "use strict";


})(window);

﻿/*
* cell-style.js
* author:devin87@qq.com
* update: 2015/11/19 17:21
*/
(function (window, undefined) {
    "use strict";
    
})(window);

﻿/*
* workbook.js
* author:devin87@qq.com
* update: 2016/03/31 15:28
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

            sheet.displayFormula = sheet.displayFormula !== false;       //是否显示公示栏
            sheet.displayGridlines = sheet.displayGridlines !== false;   //是否显示网格线
            sheet.displayHeadings = sheet.displayHeadings !== false;     //是否显示标头和行号

            sheet.isFreeze = sheet.isFreeze === true;      //是否冻结单元格
            sheet.isHidden = sheet.isHidden === true;      //Sheet是否隐藏

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

            ["firstRow", "firstCol", "lastRow", "lastCol", "scrollRow", "scrollCol", "freezeRow", "freezeCol"].forEach(function (prop) {
                sheet[prop] = sheet[prop] || 0;
            });

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

﻿/*
* ui.js
* author:devin87@qq.com
* update: 2016/04/01 17:15
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

            var html =
                '<div class="ss-toolbar"></div>' +
                '<div class="ss-formulabar"></div>' +
                '<div class="ss-sheet"></div>' +
                '<div class="ss-footbar"></div>';

            $(elView).addClass("ss-view").height(height == "auto" ? view.getHeight() : height).html(html);

            self.elView = elView;

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

﻿/*
* ui.toolbar.js 工具栏操作
* author:devin87@qq.com
* update: 2015/12/02 16:39
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

        picker = UI.picker;


    UI.extend({
        initToolbar: function () {

        }
    });

    UI.register("initToolbar");

})(window);

﻿/*
* ui.formulabar.js 公式栏操作
* author:devin87@qq.com
* update: 2015/12/02 16:39
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
        Lang = SS.Lang,
        UI = SS.UI,

        picker = UI.picker;

    
    UI.extend({
        initFormulabar: function () {

        }
    });

    UI.register("initFormulabar");

})(window);

﻿/*
* ui.cell-format.js
* author:devin87@qq.com
* update: 2016/01/15 11:20
*/
(function (window, undefined) {
    "use strict";


})(window);

﻿/*
* ui.cell-formula.js
* author:devin87@qq.com
* update: 2016/01/15 11:21
*/
(function (window, undefined) {
    "use strict";


})(window);

﻿/*
* ui.grid.js 表格操作
* update: 2016/04/13 15:20
*/
(function (window, undefined) {
    "use strict";

    var extend = Q.extend,
        async = Q.async,
        fire = Q.fire,

        isUInt = Q.isUInt,
        isUNum = Q.isUNum,

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
        //初始化表格
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

        //初始化表格大小
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

        //初始化表格事件
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

            var startX,
                startY,

                firstRow,
                firstCol,
                lastRow,
                lastCol,

                lineRow,
                lineCol,

                startWidth,
                startHeight,
                startTop,
                startLeft,
                moveX,
                moveY,

                is_mousedown,
                is_resize_row,
                is_resize_col,
                is_rightclick;

            //根据鼠标坐标获取单元格索引
            var get_cell_index = function (x, y) {
                var offset = getOffset(elSheet),
                    left = x - offset.left,
                    top = y - offset.top;

                if (left > self.FIXED_WIDTH) left += self._scrollLeft;
                if (top > self.FIXED_HEIGHT) top += self._scrollTop;

                var row = self.findRow(top),
                    col = self.findCol(left);

                if (!is_resize_row && !is_resize_col) {
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
                    startX = e.clientX;
                    startY = e.clientY;

                    var ci = get_cell_index(startX, startY);
                    lastRow = firstRow = ci.row;
                    lastCol = firstCol = ci.col;

                    if (lineRow) {
                        is_resize_row = true;
                        self.showDragHLine(lineRow);

                        startHeight = self.getRowHeight(lineRow - 1);
                        startTop = self.getRowTop(lineRow) - 1;
                    } else if (lineCol) {
                        is_resize_col = true;
                        self.showDragVLine(lineCol);

                        startWidth = self.getColWidth(lineCol - 1);
                        startLeft = self.getColLeft(lineCol) - 1;
                    } else {
                        self.selectCells(firstRow, firstCol);
                    }
                },

                mousemove: function (e) {
                    var x = e.clientX,
                        y = e.clientY,

                        ci = get_cell_index(x, y),
                        row = ci.row,
                        col = ci.col,
                        cursor = "cell";

                    if (row == INFINITY_VALUE || col == INFINITY_VALUE) {
                        cursor = "default";
                        if (lineRow) cursor = "n-resize";
                        else if (lineCol) cursor = "e-resize";
                    }

                    elMark.style.cursor = cursor;

                    if (is_resize_row) {
                        moveY = y - startY;
                        if (startHeight + moveY < 2) moveY = 2 - startHeight;

                        self.elDragHLine.style.top = (startTop + moveY) + "px";
                    } else if (is_resize_col) {
                        moveX = x - startX;
                        if (startWidth + moveX < 2) moveX = 2 - startWidth;

                        self.elDragVLine.style.left = (startLeft + moveX) + "px";
                    } else {
                        if (!is_mousedown || (row == lastRow && col == lastCol)) return;

                        lastRow = row;
                        lastCol = col;

                        self.selectCells(firstRow, firstCol, lastRow, lastCol);
                    }
                },

                mouseup: function (e) {
                    is_mousedown = false;

                    if (is_resize_row) {
                        is_resize_row = false;
                        self.hideDragHLine();

                        self.updateRowHeight(lineRow - 1, startHeight + moveY);
                    } else if (is_resize_col) {
                        is_resize_col = false;
                        self.hideDragVLine();

                        self.updateColWidth(lineCol - 1, startWidth + moveX);
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

        //更新列偏移数据
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

                offset_width += +map_col_width[key] || 0;
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
        //更新行偏移数据
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

                offset_height += +map_row_height[key] || 0;
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

        //获取列偏移
        getColLeft: function (col) {
            var self = this,
                value = self._colOffsets[col];

            if (value == undefined) {
                self._colOffsets[col] = value = calcCellOffset(self._colWidthIndexs, self._colWidthOffsets, self._colWidth, col) + self._rowNumWidth;
            }

            return value;
        },
        //获取行偏移
        getRowTop: function (row) {
            var self = this,
                value = self._rowOffsets[row];

            if (value == undefined) {
                self._rowOffsets[row] = value = calcCellOffset(self._rowHeightIndexs, self._rowHeightOffsets, self._rowHeight, row) + self._colHeadHeight;
            }

            return value;
        },

        //获取单元格元素实际左边距
        getCellLeft: function (col) {
            var self = this,
                left = self.getColLeft(col);

            if (col >= self._splitCol) left -= self.FIXED_WIDTH;

            return left;

            //return this.getColLeft(col);
        },

        //获取单元格元素实际上边距
        getCellTop: function (row) {
            var self = this,
                top = self.getRowTop(row);

            if (row >= self._splitRow) top -= self.FIXED_HEIGHT;

            return top;

            //return this.getRowTop(row);
        },

        //获取列宽
        getColWidth: function (col) {
            var w = this.sheet.colWidths[col];
            return w != undefined ? w || 0 : this._colWidth;
        },

        //获取列宽度（多列）
        getColsWidth: function (firstCol, lastCol) {
            return firstCol != lastCol ? this.getColLeft(lastCol + 1) - this.getColLeft(firstCol) : this.getColWidth(firstCol);
        },

        //获取行高
        getRowHeight: function (row) {
            var h = this.sheet.rowHeights[row];
            return h != undefined ? h || 0 : this._rowHeight;
        },

        //获取行高度（多行）
        getRowsHeight: function (firstRow, lastRow) {
            return firstRow != lastRow ? this.getRowTop(lastRow + 1) - this.getRowTop(firstRow) : this.getRowHeight(firstRow);
        },

        //根据偏移查找所知行
        findRow: function (top) {
            var self = this;

            return findCellIndex(self._rowTops, self._rowTopIndexs, self._rowHeight, top - self._colHeadHeight);
        },

        //根据偏移查找所知列
        findCol: function (left) {
            var self = this;

            return findCellIndex(self._colLefts, self._colLeftIndexs, self._colWidth, left - self._rowNumWidth);
        },

        //更改列宽
        updateColWidth: function (col, width) {
            var self = this,
                sheet = self.sheet,
                splitCol = self._splitCol;

            if (!isUInt(col) || !isUNum(width)) return self;

            sheet.colWidths[col] = width;

            self.updateColsOffset();

            if (col < splitCol) self.updateGridSize().showSplitVLine(splitCol);

            var updateEl = function (list) {
                var el = list[col];
                if (el) {
                    el.style.width = width + "px";
                    getFirst(el).style.width = (width - 1) + "px";
                }

                for (var i = col + 1, len = list.length; i < len; i++) {
                    el = list[i];
                    if (el) el.style.left = self.getCellLeft(i) + "px";
                }
            };

            updateEl(self.elColHeads);
            self.elCells.forEach(updateEl);

            return self.drawMerges().selectCellsAuto().updateHScrollWidth();
        },

        //更改行高
        updateRowHeight: function (row, height) {
            var self = this,
                sheet = self.sheet,
                splitRow = self._splitRow;

            if (!isUInt(row) || !isUNum(height)) return self;

            sheet.rowHeights[row] = height;

            self.updateRowsOffset();

            if (row < splitRow) self.updateGridSize().showSplitHLine(splitRow);

            var elRowNums = self.elRowNums,
                elCells = self.elCells;

            for (var i = row, len = elRowNums.length; i < len; i++) {
                var elSSRowLeft = self.getElSSRowLeft(i),
                    elSSRowRight = self.getElSSRowRight(i);

                if (i == row) {
                    if (elSSRowLeft) elSSRowRight.style.height = elSSRowLeft.style.height = height + "px";

                    var cssInnerHeight = (height - 1) + "px";
                    if (elRowNums[i]) getFirst(elRowNums[i]).style.height = cssInnerHeight;
                    elCells[i].forEach(function (el, j) {
                        if (el) getFirst(el).style.height = cssInnerHeight;
                    });
                } else {
                    if (elSSRowLeft) elSSRowRight.style.top = elSSRowLeft.style.top = self.getCellTop(i) + "px";
                }
            }

            return self.drawMerges().selectCellsAuto().updateVScrollHeight();
        },

        //获取所有单元格宽度
        getTotalWidth: function () {
            return this.getColsWidth(0, this._maxCol + 1);
        },

        //获取所有单元格高度
        getTotalHeight: function () {
            return this.getRowsHeight(0, this._maxRow + 1);
        },

        //获取水平滚动条宽度
        getHScrollWidth: function () {
            var self = this;
            return self._HScrollbarWidth + self.getColsWidth(self._splitCol, self._maxCol + 1) - self.SCROLL_WIDTH;
        },

        //获取垂直滚动条高度
        getVScrollHeight: function () {
            var self = this;
            return self._VScrollbarHeight + self.getRowsHeight(self._splitRow, self._maxRow + 1) - self.SCROLL_HEIGHT;
        },

        //更新水平滚动条宽度
        updateHScrollWidth: function () {
            var self = this,
                elHScroll = self.elHScroll;

            if (elHScroll) getFirst(elHScroll).style.width = self.getHScrollWidth() + "px";

            return self;
        },

        //更新垂直滚动条高度
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

        //更新表格设置
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

        //更新合并单元格数据
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

        //获取合并单元格数据
        getMergeData: function (row, col) {
            //INFINITY_VALUE 表示整行或整列
            var mapMerge = this.mapMerge,
                mapRow = mapMerge[INFINITY_VALUE] || mapMerge[row];

            return mapRow ? mapRow[INFINITY_VALUE] || mapRow[col] : undefined;
        },

        //获取合并状态
        //RANGE_NO_MERGE : 非合并区域
        //RANGE_IS_MERGE : 合并区域
        //RANGE_IS_MERGE_FIRST: 合并区域中的第一个单元格
        //RANGE_IS_MERGE_OTHER: 合并区域中的其它单元格
        getMergeState: function (firstRow, firstCol, lastRow, lastCol) {
            var self = this,
                data = self.getMergeData(firstRow, firstCol);

            if (!data) return RANGE_NO_MERGE;

            var isFirst = data[0] == firstRow && data[1] == firstCol;
            if (isFirst && data[2] == lastRow && data[3] == lastCol) return RANGE_IS_MERGE;

            return isFirst ? RANGE_IS_MERGE_FIRST : RANGE_IS_MERGE_OTHER;
        },

        //是否是合并单元格
        isMergeCell: function (firstRow, firstCol, lastRow, lastCol) {
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

        //获取合并单元格元素
        getElMerges: function (data) {
            return self.mapElMerges[data.join("_")];
        },

        //合并单元格 data => [firstRow, firstCol, lastRow, lastCol]
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

        //合并单元格
        mergeCells: function (firstRow, firstCol, lastRow, lastCol) {
            return this._mergeCells([firstRow, firstCol, lastRow, lastCol]);
        },

        //绘制合并单元格
        drawMerges: function (list) {
            var self = this;
            self.elMerges4.innerHTML = self.elMerges3.innerHTML = self.elMerges2.innerHTML = self.elMerges1.innerHTML = "";

            self.mapElMerges = {};
            (list || self.sheet.merges).forEach(function (data) {
                self._mergeCells(data);
            });

            return self;
        },

        //初始化合并单元格
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

            return self.updateGridMerges().drawMerges();
        },

        //更新表格大小
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

        //获取行号元素(.ss-rowNum)
        getElRowNum: function (row) {
            return this.elRowNums[row];
        },
        //获取列头元素(.ss-colHead)
        getElColHead: function (col) {
            return this.elColHeads[col];
        },
        //获取整行单元格元素(.ss-cell)
        getElRowCells: function (row) {
            return this.elCells[row];
        },
        //获取单元格元素(.ss-cell)
        getElCell: function (row, col) {
            var elRowCells = this.getElRowCells(row);
            return elRowCells ? elRowCells[col] : undefined;
        },

        //获取左边行元素(.ss-grid3 .ss-row)
        getElSSRowLeft: function (row) {
            var elRowNum = this.getElRowNum(row);
            if (elRowNum) return elRowNum.parentNode;
        },
        //获取右边行元素(.ss-grid4 .ss-row)
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

        //绘制行
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

        //绘制列
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

        //更新元素缓存
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

        //设置行号
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

        //滚动到行
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
        //滚动到列
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

        //滚动到行（异步）
        scrollToRowAsync: function (row, isUpdateScrollbar, callback) {
            var self = this;

            async(function () {
                self.scrollToRow(row, isUpdateScrollbar);
                fire(callback, self);
            }, 0);

            return self;
        },
        //滚动到列（异步）
        scrollToColAsync: function (col, isUpdateScrollbar, callback) {
            var self = this;

            async(function () {
                self.scrollToCol(col, isUpdateScrollbar);
                fire(callback, self);
            }, 0);

            return self;
        },

        //选中列高亮
        activeColHeads: function (firstCol, lastCol) {
            var self = this;

            $(self.elColHeads.slice(firstCol, lastCol + 1)).addClass("ss-on");

            return self;
        },

        //选中行高亮
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

        //显示行分隔线
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
        //显示列分隔线
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

        //显示分隔线（行列）
        showSplitLine: function () {
            var self = this;

            return self.showSplitHLine(self._splitRow).showSplitVLine(self._splitCol);
        },

        //显示行拖拽线
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
        //显示列拖拽线
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

        //隐藏行拖拽线
        hideDragHLine: function () {
            var self = this,
                elDragHLine = self.elDragHLine;

            if (elDragHLine) elDragHLine.style.display = "none";

            return self;
        },
        //隐藏列拖拽线
        hideDragVLine: function () {
            var self = this,
                elDragVLine = self.elDragVLine;

            if (elDragVLine) elDragVLine.style.display = "none";

            return self;
        },

        //显示选择面板
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

        //选择单元格并显示选择面板
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

                isSelectOne = self.isMergeCell(firstRow, firstCol, lastRow, lastCol);
            }

            self._isSelectOne = isSelectOne;

            self.showPanel(firstRow, firstCol, lastRow, lastCol);

            sheet.firstRow = firstRow;
            sheet.firstCol = firstCol;
            sheet.lastRow = lastRow;
            sheet.lastCol = lastCol;

            return self;
        },

        //选择单元格
        selectCellsAuto: function () {
            var self = this,
                sheet = self.sheet;

            return self.selectCells(sheet.firstRow, sheet.firstCol, sheet.lastRow, sheet.lastCol);
        },

        //绘制单元格
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

        //绘制表格
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

﻿/*
* ui.footbar.js 底部栏操作
* author:devin87@qq.com
* update: 2016/04/05 17:24
*/
(function (window, undefined) {
    "use strict";

    var extend = Q.extend,
        vals = Q.vals,
        toMap = Q.toMap,

        getFirst = Q.getFirst,
        createEle = Q.createEle,

        measureText = Q.getMeasureText(),
        view = Q.view,

        DragX = Q.DragX,
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

            //水平滚动条宽高
            self._HScrollbarWidth = elHScroll.offsetWidth;
            self._HScrollbarHeight = elHScroll.offsetHeight;

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

            //底部拖动更改水平滚动条宽度事件
            new DragX(function () {
                var base = this,
                    minWidth = 200,
                    maxWidth = 0,
                    currentWidth = 0,
                    startX;

                //必须,拖动元素
                base.ops = { ele: elResize, autoCss: false, autoIndex: false };

                //处理函数:鼠标按下时触发
                base.doDown = function (e) {
                    startX = e.clientX;

                    currentWidth = elHScroll.offsetWidth;
                    maxWidth = self.GRID_WIDTH - 60 - getFirst(elTabs).offsetWidth;
                };

                //处理函数:拖动时触发
                base.doMove = function (e) {
                    var x = e.clientX - startX;
                    if (x == 0) return;

                    var width = currentWidth - x;
                    if (width > maxWidth) width = maxWidth;
                    else if (width < minWidth) width = minWidth;

                    elHScroll.style.width = width + "px";
                };

                //处理函数:鼠标释放时触发
                base.doUp = function () {
                    //for 通用处理
                    self.setHScrollbarWidth(elHScroll.offsetWidth);
                };
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

            //var html =
            //    sheets.map(function (sheet) {
            //        if (!workbook.isDisplaySheet(sheet)) return '';

            //        var index = sheet.index,
            //            name = sheet.name,
            //            tabColor = sheet.tabColor;

            //        var html =
            //            '<div class="sst-item' + (index == sheetIndex ? ' sst-on' : '') + '" title="' + name + '"' + (tabColor ? ' style="color:' + tabColor + ';"' : '') + '>' +
            //                '<div class="sst-name">' + name + '</div>' +
            //            '</div>';

            //        return html;
            //    }).join('') +
            //    '<div class="sst-insert" title="' + Lang.INSERT_SHEET + '">' +
            //        '<div class="sst-name">+</div>' +
            //    '</div>';

            var html =
                '<table class="ss-tabs">' +
                    '<tr>' +
                        sheets.map(function (sheet) {
                            if (!workbook.isDisplaySheet(sheet)) return '';

                            var index = sheet.index,
                                name = sheet.name,
                                tabColor = sheet.tabColor;

                            var html =
                                '<td class="sst-item' + (index == sheetIndex ? ' sst-on' : '') + '" title="' + name + '"' + (tabColor ? ' style="color:' + tabColor + ';"' : '') + '>' +
                                    '<div class="sst-name">' + name + '</div>' +
                                '</td>';

                            return html;
                        }).join('') +
                        '<td class="sst-insert" title="' + Lang.INSERT_SHEET + '">' +
                            '<div class="sst-name">+</div>' +
                        '</td>' +
                    '</tr>' +
                '</table>';

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

            var elHScroll = self.elHScroll,
                scrollbarWidth = elHScroll.offsetWidth,
                maxScrollbarWidth = self.GRID_WIDTH - 60 - getFirst(elTabs).offsetWidth;

            if (scrollbarWidth > maxScrollbarWidth) self.setHScrollbarWidth(maxScrollbarWidth);
        },
        //设置水平滚动条宽度
        setHScrollbarWidth: function (width) {
            var self = this;
            self._HScrollbarWidth = width;
            self.elHScroll.style.width = width + "px";
            return self;
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

﻿/*
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