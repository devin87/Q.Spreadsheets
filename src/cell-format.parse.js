/// <reference path="core.js" />
/*
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