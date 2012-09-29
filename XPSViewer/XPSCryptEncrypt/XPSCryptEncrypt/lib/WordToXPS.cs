using System;
using System.Net;
//using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Documents;
//using System.Windows.Ink;
//using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Animation;
//using System.Windows.Shapes;

using Word = Microsoft.Office.Interop.Word;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace XPSCryptEncrypt.Lib
{
    public class WordToXPS : Iconverter
    {
        public Word.WdExportFormat exportFormat = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatXPS;

        public  bool Convert(string sourcePath, string targetPath)
        {
            bool result;
            object paramMissing = Type.Missing;
            Word.Application wordApplication = new Word.Application();
            //Word.ApplicationClass wordApplication = new Word.Application();
            Word.Document wordDocument = null;
            try
            {
                object paramSourceDocPath = sourcePath;
                string paramExportFilePath = targetPath;

                Word.WdExportFormat paramExportFormat = exportFormat;
                bool paramOpenAfterExport = false;
                Word.WdExportOptimizeFor paramExportOptimizeFor = Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                Word.WdExportRange paramExportRange = Word.WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                Word.WdExportItem paramExportItem = Word.WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                Word.WdExportCreateBookmarks paramCreateBookmarks = Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;

                wordDocument = wordApplication.Documents.Open(
                        ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing);

                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                            paramExportFormat, paramOpenAfterExport,
                            paramExportOptimizeFor, paramExportRange, paramStartPage,
                            paramEndPage, paramExportItem, paramIncludeDocProps,
                            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                            paramBitmapMissingFonts, paramUseISO19005_1,
                            ref paramMissing);
                result = true;
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        //public void PrintWord(string wordfile)
        //{
        //    oWord.ApplicationClass word = new oWord.ApplicationClass();
        //    Type wordType = word.GetType();

        //    //打开WORD文档 
        //    oWord.Documents docs = word.Documents;
        //    Type docsType = docs.GetType();
        //    object objDocName = wordfile;
        //    oWord.Document doc = (oWord.Document)docsType.InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { objDocName, true, true });

        //    //打印输出到指定文件 
        //    //可以使用 doc.PrintOut();方法,次方法调用中的参数设置较繁琐,建议使用 Type.InvokeMember 来调用时可以不用将PrintOut的参数设置全,只设置4个主要参数 
        //    Type docType = doc.GetType();
        //    object printFileName = wordfile + ".xps";
        //    docType.InvokeMember("PrintOut", System.Reflection.BindingFlags.InvokeMethod, null, doc, new object[] { false, false, oWord.WdPrintOutRange.wdPrintAllDocument, printFileName });

        //    //退出WORD 
        //    wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, word, null);
        //}

    }
}
