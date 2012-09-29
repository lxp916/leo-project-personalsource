using System;
using System.Net;
using Microsoft.Office.Interop.Excel;

namespace XPSCryptEncrypt.Lib
{
    public class ExcelToXPS:Iconverter
    {
        XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypeXPS;
        public bool Convert(string sourcePath, string targetPath )
        {
            bool result;
            object missing = Type.Missing;
            Application application = null;
            Workbook workBook = null;
            try
            {
                application = new Application();
                object target = targetPath;
                object type = targetType;
                workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);

                workBook.ExportAsFixedFormat(targetType, target, XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }


        ///// <summary>
        ///// Excel利用虚拟打印机转换为xps
        ///// </summary>
        ///// <param name="execlfile"></param>
        //        public void PrintExcel(string execlfile) 
        //        { 
        //                Excel.ApplicationClass eapp = new Excel.ApplicationClass(); 
        //                Type eType = eapp.GetType(); 
        //                Excel.Workbooks Ewb = eapp.Workbooks; 
        //                Type elType = Ewb.GetType(); 
        //                object objelName = execlfile; 
        //                Excel.Workbook ebook = (Excel.Workbook)elType.InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, Ewb, new Object[] { objelName, true, true }); 
 
        //                object printFileName = execlfile + ".xps"; 
 
        //                Object oMissing = System.Reflection.Missing.Value; 
        //                ebook.PrintOut(oMissing, oMissing, oMissing, oMissing, oMissing, true, oMissing, printFileName); 
 
        //                eType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, eapp, null); 
        //        }
    }
}
