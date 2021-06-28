using System;
using System.Runtime.InteropServices;

using VBDE = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;


namespace watermark
{
    class ExcelOption
    {
        // 对指定的excel文件插入宏并另存为
        public void InsertMacro(string srcFile, string dstFile, string macroString)
        {
            Excel.Application xl = null;
            Excel._Workbook wb = null;
            VBDE.VBComponent module = null;
            bool SaveChanges = false;

            try
            {
                // 打开excel文件
                xl = new Excel.Application();
                xl.DisplayAlerts = false;
                wb = xl.Workbooks.Open(srcFile);
                wb.CheckCompatibility = false;
                wb.DoNotPromptForConvert = true;

                // 插入宏
                module = wb.VBProject.VBComponents.Add(VBDE.vbext_ComponentType.vbext_ct_StdModule);
                module.CodeModule.AddFromString(macroString);

                xl.Visible = false;
                xl.UserControl = false;
                SaveChanges = true;

                // 保存文件
                //wb.SaveAs(dstFile, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
                //    null, null, false, false, Excel.XlSaveAsAccessMode.xlShared,
                //    false, false, null, null, null);
                wb.SaveAs(dstFile, Excel.XlFileFormat.xlWorkbookNormal,
                    null, null, false, false, Excel.XlSaveAsAccessMode.xlShared,
                    false, false, null, null, null);
            }
            catch (Exception theException)
            {
                String msg;
                msg = "Error: ";
                msg = String.Concat(msg, theException.Message);
                msg = String.Concat(msg, " Line: ");
                msg = String.Concat(msg, theException.Source);
                Console.WriteLine(msg);
                throw theException;
            }
            finally
            {
                // 释放资源
                try
                {
                    xl.Visible = false;
                    xl.UserControl = false;
                    wb.Close(SaveChanges, null, null);
                    xl.Workbooks.Close();
                }
                catch { }

                xl.Quit();

                if (module != null) { Marshal.ReleaseComObject(module); }
                if (wb != null) { Marshal.ReleaseComObject(wb); }
                if (xl != null) { Marshal.ReleaseComObject(xl); }

                module = null;
                wb = null;
                xl = null;
                GC.Collect();
            }
        }

        //// 创建一个空的excel文件
        //public void CreateEmptyWorkbook(string FileName)
        //{ 

        //}

        //public void CreateWorkbook(string FileName, string Macro)
        //{

        //    Excel.Application xl = null;
        //    Excel._Workbook wb = null;
        //    Excel._Worksheet sheet = null;
        //    VBDE.VBComponent module = null;
        //    bool SaveChanges = false;


        //    try
        //    {

        //        if (File.Exists(FileName)) { File.Delete(FileName); }

        //        GC.Collect();

        //        xl = new Excel.Application
        //        {
        //            Visible = false
        //        };

        //        wb = xl.Workbooks.Add(Missing.Value);
        //        sheet = (Excel._Worksheet)wb.ActiveSheet;

        //        module = wb.VBProject.VBComponents.Add(VBDE.vbext_ComponentType.vbext_ct_StdModule);
        //        module.CodeModule.AddFromString(Macro);

        //        xl.Visible = false;
        //        xl.UserControl = false;
        //        SaveChanges = true;

        //        wb.SaveAs(FileName, Excel.XlFileFormat.xlWorkbookNormal,
        //        null, null, false, false, Excel.XlSaveAsAccessMode.xlShared,
        //        false, false, null, null, null);

        //    }
        //    catch (Exception theException)
        //    {
        //        String msg;
        //        msg = "Error: ";
        //        msg = String.Concat(msg, theException.Message);
        //        msg = String.Concat(msg, " Line: ");
        //        msg = String.Concat(msg, theException.Source);
        //        Console.WriteLine(msg);
        //    }
        //    finally
        //    {

        //        try
        //        {
        //            xl.Visible = false;
        //            xl.UserControl = false;
        //            wb.Close(SaveChanges, null, null);
        //            xl.Workbooks.Close();
        //        }
        //        catch { }

        //        xl.Quit();

        //        if (module != null) { Marshal.ReleaseComObject(module); }
        //        if (sheet != null) { Marshal.ReleaseComObject(sheet); }
        //        if (wb != null) { Marshal.ReleaseComObject(wb); }
        //        if (xl != null) { Marshal.ReleaseComObject(xl); }

        //        module = null;
        //        sheet = null;
        //        wb = null;
        //        xl = null;
        //        GC.Collect();
        //    }
        //}
    }
}
