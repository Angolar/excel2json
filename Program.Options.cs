using System;
using CommandLine;
using CommandLine.Text;

namespace excel2json
{
    partial class Program
    {
        /// <summary>
        /// 命令行参数定义
        /// </summary>
        internal sealed class Options
        {
            public Options()
            {
                this.HeaderRows = 3;
                this.Encoding = "utf8-nobom";
                this.Lowcase = false;
                this.ExportArray = false;
                this.ForceSheetName = false;
            }

            /// <summary>
            /// -e, –excel Required. 输入的Excel文件路径.
            /// </summary>
            [Option('e', "excel", Required = true, HelpText = "input excel file path.")]
            public string ExcelPath {
                get;
                set;
            }

            /// <summary>
            /// -j, –json 指定输出的json文件路径.
            /// </summary>
            [Option('j', "json", Required = false, HelpText = "export json file path.")]
            public string JsonPath {
                get;
                set;
            }

            /// <summary>
            /// -p, –csharp 指定输出的C#文件路径.
            /// </summary>
            [Option('p', "csharp", Required = false, HelpText = "export C# data struct code file path.")]
            public string CSharpPath {
                get;
                set;
            }

            /// <summary>
            /// -h, –header (Default: 3)表格中有几行是表头.
            /// </summary>
            [Option('h', "header", Required = false, DefaultValue = 1, HelpText = "number lines in sheet as header.")]
            public int HeaderRows {
                get;
                set;
            }

            /// <summary>
            /// -c, –encoding (Default: utf8-nobom) 指定编码的名称.
            /// </summary>
            [Option('c', "encoding", Required = false, DefaultValue = "utf8-nobom", HelpText = "export file encoding.")]
            public string Encoding {
                get;
                set;
            }

            /// <summary>
            /// -l, –lowcase (Default: false) 自动把字段名称转换成小写格式.
            /// </summary>
            [Option('l', "lowcase", Required = false, DefaultValue = false, HelpText = "convert filed name to lowcase.")]
            public bool Lowcase {
                get;
                set;
            }

            /// <summary>
            /// -a 序列化成数组
            /// </summary>
            [Option('a', "array", Required = false, DefaultValue = false, HelpText = "export as array, otherwise as dict object.")]
            public bool ExportArray {
                get;
                set;
            }

            /// <summary>
            /// -d, –date:指定日期格式化字符串，例如：dd / MM / yyy hh: mm:ss
            /// </summary>
            [Option('d', "date", Required = false, DefaultValue = "yyyy/MM/dd", HelpText = "Date Format String, example: dd / MM / yyy hh: mm:ss.")]
            public string DateFormat {
                get;
                set;
            }

            /// <summary>
            /// -s 序列化时强制带上sheet name，即使只有一个sheet
            /// </summary>
            [Option('s', "sheet", Required = false, DefaultValue = false, HelpText = "export with sheet name, even there's only one sheet.")]
            public bool ForceSheetName {
                get;
                set;
            }

            /// <summary>
            /// -exclude_prefix： 导出时，排除掉包含指定前缀的表单和列，例如：-exclude_prefix #
            /// </summary>
            [Option('x', "exclude_prefix", Required = false, DefaultValue = "", HelpText = "exclude sheet or column start with specified prefix.")]
            public string ExcludePrefix {
                get;
                set;
            }

            /// <summary>
            /// -cell_json：自动识别单元格中的Json对象和Json数组，Default：false
            /// </summary>
            [Option('l', "cell_json", Required = false, DefaultValue = false, HelpText = "convert json string in cell")]
            public bool CellJson {
                get;
                set;
            }

            [Option('l', "all_string", Required = false, DefaultValue = false, HelpText = "all string")]
            public bool AllString
            {
                get;
                set;
            }
        }
    }
}
