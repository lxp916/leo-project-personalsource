using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace XPSCryptEncrypt.Lib
{
    public class Common
    {
        public enum FileFormat
        {
            PDF,Word,Excel,PowerPoint,Image,NULL
        }

        public static FileFormat getFileFormat(string extension)
        {
            FileFormat format = FileFormat.NULL;
            extension= extension.ToLower();
            if (extension == "doc" || extension == "docx")
                format = FileFormat.Word;
            else if (extension == "xls" || extension == "xlsx")
                format = FileFormat.Excel;
            else if (extension == "ppt" || extension == "pptx")
                format = FileFormat.PowerPoint;
            else if (extension == "jpg" || extension == "png")
                format = FileFormat.Image;
            else if (extension == "pdf" )
                format = FileFormat.PDF;
            return format;
        }

    }
}
