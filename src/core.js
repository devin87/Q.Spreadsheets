/// <reference path="util.js" />
/*
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