using System;
using System.Net;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;


namespace XPSCryptEncrypt.Lib
{
    public class PowerPointToXPS : Iconverter
    {
        PpSaveAsFileType targetFileType = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsXPS;
        public bool Convert(string sourcePath, string targetPath)
        {
            bool result;
            object missing = Type.Missing;
            Application application = null;
            Presentation persentation = null;
            try
            {
                application = new Application();
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue);

                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
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
    }
}
