using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XPSCryptEncrypt.Lib
{
    public class ConverterProxy : Iconverter
    {
        Iconverter proxy;
        private ConverterProxy() { }
        public ConverterProxy(string fileExtension) 
        {
            switch (Common.getFileFormat(fileExtension))
            {
                case Common.FileFormat.Word:
                    proxy = new WordToXPS();
                    break;
                case Common.FileFormat.Excel:
                    proxy = new ExcelToXPS();
                    break;
                case Common.FileFormat.PowerPoint:
                    proxy = new PowerPointToXPS();
                    break;
                case Common.FileFormat.Image:
                    proxy = new ImageToXPS();
                    break;
                case Common.FileFormat.PDF:
                    proxy = new PDFToXPS();
                    break;
            }
        }

        public bool Convert(string sourcePath, string targetPath)
        {
            return object.Equals(proxy,null) ? false:this.proxy.Convert(sourcePath, targetPath);
        }
    }
}