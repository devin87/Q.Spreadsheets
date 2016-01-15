//build 配置文件
module.exports = {
    root: "../",
    output: "dist",

    noStore: true,

    concat: {
        title: "文件合并",

        list: [
            {
                dir: "src",
                src: ["init.js", "util.js", "core.js", "events.js", "errors.js", "cell-format.parse.js", "cell-format.js", "cell-formula.parse.js", "cell-formula.js", "cell-style.js", "workbook.js", "ui.js", "ui.toolbar.js", "ui.formulabar.js", "ui.cell-format.js", "ui.cell-formula.js", "ui.grid.js", "ui.footbar.js", "spreadsheets.js"],
                dest: "/Q.Spreadsheets.js",
                prefix: "/* Web Spreadsheets, https://github.com/devin87/Q.Spreadsheets */\n"
            },
            {
                dir: "lib",
                src: ["Q.js", "Q.UI.js"],
                dest: "/Q.all.js"
            }
        ],

        replace: [
            //移除\r字符
            [/\r/g, ""],
            //移除VS引用
            [/\/\/\/\s*<reference path="[^"]*" \/>\n/gi, ""]
        ]
    },

    cmd: [
        {
            title: "压缩js",

            //cmd: "java -jar D:\\tools\\compiler.jar --js=%f.fullname% --js_output_file=%f.dest%",
            cmd: "uglifyjs %f.fullname% -o %f.dest% -c -m",

            match: ["*.js"],

            replace: [
                //去掉文件头部压缩工具可能保留的注释
                [/^\/\*([^~]|~)*?\*\//, ""]
            ],

            //可针对单一的文件配置 before、after,def 表示默认
            before: [
                {
                    "def": "//devin87@qq.com\n",
                    "Q.js": "//Q.js devin87@qq.com\n//mojoQuery&mojoFx scott.cgi\n",
                    "Q.UI.js": "//Q.UI.js devin87@qq.com\n",
                    "jquery-1.11.3.js": "/*! jQuery v1.11.3 | (c) 2005, 2015 jQuery Foundation, Inc. | jquery.org/license */\n"
                },
                "//build:%NOW%\n"
            ]
        }
    ],

    copy: [
        {
            title: "同步lang",
            dir: "src/lang",
            output: "lang",
            match: "*.js"
        }
    ],

    run: ["concat", "cmd", "copy"]
};