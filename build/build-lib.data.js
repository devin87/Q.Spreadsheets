//build 配置文件
module.exports = {
    root: "../",

    noStore: true,

    concat: {
        title: "文件合并",

        list: [
            {

                dir: "../Q.js/src",
                src: ["Q.js", "Q.Queue.js", "Q.core.js", "Q.setTimer.js", "Q.query.js", "Q.dom.js", "Q.event.js", "Q.ajax.js", "Q.$.js", "Q.animate.js"],
                dest: "/lib/Q.js"
            },
            {
                dir: "../Q.UI.js/src",
                src: ["Q.UI.Box.js", "Q.UI.ContextMenu.js", "Q.UI.DropdownList.js", "Q.UI.DataPager.js", "Q.UI.ColorPicker.js"],
                dest: "/lib/Q.UI.js"
            }
        ],

        replace: [
            //移除\r字符
            [/\r/g, ""],
            //移除VS引用
            [/\/\/\/\s*<reference path="[^"]*" \/>\n/gi, ""]
        ]
    },

    run: ["concat"]
};