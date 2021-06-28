using System;
using System.IO;
using System.Linq;

namespace watermark
{
    class Program
    {
        // args:[SrcFile, DstFile, VbaScript]
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                // 参数个数不足
                Console.WriteLine("参数个数不足");
                Environment.Exit(1);
            }
            string SrcFile = args[0];
            string DstFile = args[1];
            string VbaScript = args[2];

            string[] InExcelSupport = new string[] { ".xls", ".xlsm", ".xlsx" }; // excel文件支持列表
            string[] OutExcelSupport = new string[] { ".xlsm", ".xls" }; // excel 文件输出列表

            string[] InWordSupport = new string[] { ".doc", ".docm", ".docx" }; // word文件支持列表
            string[] OutWordSupport = new string[] { ".doc", ".docm" }; // word文件输出支持列表

            // 判断源文件地址与目的文件地址是否存在
            if (!File.Exists(SrcFile) || File.Exists(DstFile) || !File.Exists(VbaScript))
            {
                Console.WriteLine(!File.Exists(SrcFile));
                Console.WriteLine(File.Exists(DstFile));
                Console.WriteLine(!File.Exists(VbaScript));

                Console.WriteLine("源文件文件已存在或目的文件存在");
                Environment.Exit(2);
            }

            // 判断文件类型
            string SrcFileSuffix = Path.GetExtension(SrcFile);
            string DstFileSuffix = Path.GetExtension(DstFile);
            Program program = new Program();
            if (InExcelSupport.Contains(SrcFileSuffix) && OutExcelSupport.Contains(DstFileSuffix))
            {
                // excel宏插入
                try
                {
                    program.ExcelMacroInsert(SrcFile, DstFile, VbaScript);
                }
                catch (Exception theException)
                {
                    Console.WriteLine(theException.ToString());
                    Environment.Exit(4);
                }
                Console.WriteLine("excel 宏插入完成");
            }
            else if (InWordSupport.Contains(SrcFileSuffix) && OutWordSupport.Contains(DstFileSuffix))
            {
                // word 宏插入
                try
                {
                    program.WordMacroInsert(SrcFile, DstFile, VbaScript);
                }
                catch (Exception theException)
                {
                    Console.WriteLine(theException.ToString());
                    Environment.Exit(5);
                }
                Console.WriteLine("word 宏插入完成");
            }
            else
            {
                // 输入输出文件类型异常
                Console.WriteLine("输入输出文件类型异常");
                Environment.Exit(3);
            }

            // word宏插入示例
            //string srcFile = Path.GetFullPath(".") + @"\test.docx";
            //string dstFile = Path.GetFullPath(".") + @"\dst\test1.docm";
            //string sCode =
            //     "sub AutoOpen()\r\n" +
            //     "Application.DisplayAlerts = False \r\n" +
            //     "   msgbox \"VBA Macro called\"\r\n" +
            //     "Application.DisplayAlerts = True \r\n" +
            //     "end sub";

            //var wordOption = new WordOption();
            //wordOption.InsertMacro(srcFile, dstFile, sCode);
            //Console.WriteLine("Word 生成完成");
            //Console.ReadLine();

            // excel宏插入示例
            //string srcFile = Path.GetFullPath(".") + @"\test.xlsx";
            //string dstFile = Path.GetFullPath(".") + @"\dst\test.xlsm";
            //string sCode =
            //     "sub Auto_Open()\r\n" +
            //     "Application.DisplayAlerts = False \r\n" +
            //     "   msgbox \"VBA Macro called\"\r\n" +
            //     "Application.DisplayAlerts = True \r\n" +
            //     "end sub";
            //var excelOption = new ExcelOption();
            //excelOption.InsertMacro(srcFile, dstFile, sCode);
            //Console.WriteLine("Excel 生成完成");
            //Console.ReadLine();
        }

        // word类型宏插入
        public void WordMacroInsert(string SrcFile, string DstFile, string VbaScript)
        {
            string sCode = File.ReadAllText(VbaScript);
            var wordOption = new WordOption();
            wordOption.InsertMacro(SrcFile, DstFile, sCode);
        }

        // excel 类型宏插入
        public void ExcelMacroInsert(string SrcFile, string DstFile, string VbaScript)
        {
            string sCode = File.ReadAllText(VbaScript);
            var excelOption = new ExcelOption();
            excelOption.InsertMacro(SrcFile, DstFile, sCode);
        }
    }
}
