using System;

using Microsoft.Vbe.Interop;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace watermark
{
    class WordOption
    {
        public void InsertMacro(string SrcFile, string DstFile, string Macro)
        {
            Word.Application oWord = null;
            Document oDoc = null;
            VBComponent oModule = null;
            bool SaveChanges = false;

            try
            {
                // 打开word文档
                oWord = new Word.Application();
                oDoc = oWord.Documents.Open(SrcFile);

                // 插入宏
                oModule = oDoc.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                oModule.CodeModule.AddFromString(Macro);
                SaveChanges = true;

                // 保存文件
                oWord.Visible = false;
                oDoc.UserControl = false;
                //oDoc.SaveAs2(DstFile, WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode: WdCompatibilityMode.wdWord2010);
                oDoc.SaveAs2(DstFile, WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode: WdCompatibilityMode.wdWord2010);
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
                // 资源释放
                try
                {
                    oWord.Visible = false;
                    oDoc.UserControl = false;
                    oDoc.Close(SaveChanges, null, null);
                    oWord.Documents.Close();
                }
                catch { }

                oWord.Quit();

                if (oModule != null) { Marshal.ReleaseComObject(oModule); }
                if (oDoc != null) { Marshal.ReleaseComObject(oDoc); }
                if (oWord != null) { Marshal.ReleaseComObject(oWord); }

                oModule = null;
                oDoc = null;
                oWord = null;
                GC.Collect();
            }
        }

        //public void InsertMacro(string srcFile, string dstFile, string macroString)
        //{
        //    Microsoft.Office.Interop.Word.Application oWord;
        //    Microsoft.Office.Interop.Word.Document oDoc;
        //    Office.CommandBar oCommandBar;
        //    Office.CommandBarButton oCommandBarButton;
        //    String sCode;
        //    Object oMissing = System.Reflection.Missing.Value;

        //    oWord = new Microsoft.Office.Interop.Word.Application();
        //    oDoc = oWord.Documents.Open(srcFile);
        //    //oDoc = oWord.Documents.Add(oMissing);
        //    try
        //    {
        //        // Create a new VBA code module.
        //        var oModule = oDoc.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
        //        sCode =
        //        "sub AutoOpen()\r\n" +
        //        "Application.DisplayAlerts = False \r\n" +
        //        "   msgbox \"VBA Macro called\"\r\n" +
        //        "Application.DisplayAlerts = True \r\n" +
        //        "end sub";
        //        // Add the VBA macro to the new code module.
        //        oModule.CodeModule.AddFromString(sCode);
        //        oModule = null;
        //    }
        //    catch (Exception e)
        //    {
        //        if (e.ToString().Contains("不被信任"))
        //            Console.WriteLine("到 Visual Basic Project 的程序访问不被信任", "Error");
        //        return;
        //    }
        //    try
        //    {
        //        // Create a new toolbar and show it to the user.
        //        oCommandBar = oWord.CommandBars.Add("VBAMacroCommandBar", oMissing, oMissing);
        //        oCommandBar.Visible = true;
        //        // Create a new button on the toolbar.
        //        oCommandBarButton = (Office.CommandBarButton)oCommandBar.Controls.Add(
        //        Office.MsoControlType.msoControlButton,
        //        oMissing, oMissing, oMissing, oMissing);
        //        // Assign a macro to the button.
        //        oCommandBarButton.OnAction = "VBAMacro";
        //        // Set the caption of the button.
        //        oCommandBarButton.Caption = "Call VBAMacro";
        //        // Set the icon on the button to a picture.
        //        oCommandBarButton.FaceId = 2151;
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("VBA宏命令已经存在.", "Error");
        //    }

        //    oDoc.SaveAs2(dstFile, WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode: WdCompatibilityMode.wdWord2010);
        //    oWord.Documents.Save();
        //    oDoc.Close();
        //    //oWord.Visible = true;

        //    oCommandBarButton = null;
        //    oCommandBar = null;
        //    oDoc = null;
        //    oWord = null;
        //    GC.Collect();


        //    Console.ReadLine();
        //}
    }
}
