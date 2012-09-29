using System;
using System.Net;
using System.Drawing;
using System.IO;
using System.Windows.Xps.Packaging;

namespace XPSCryptEncrypt.Lib
{
    public class PDFToXPS : Iconverter
    {
        public bool Convert(string sourcePath, string targetPath)
        {
            //Image img = Image.FromFile(sourcePath);

            //string resFile = sourcePath;
            //XpsDocument xpsDoc = new XpsDocument(targetPath, FileAccess.ReadWrite);
            //IXpsFixedDocumentSequenceWriter fds = xpsDoc.AddFixedDocumentSequence();
            //IXpsFixedDocumentWriter fd = fds.AddFixedDocument();
            //IXpsFixedPageWriter fp = fd.AddFixedPage();
            //XpsResource res = null;
            //XpsResource thumb = null;
            
            //res = fp.AddImage(XpsImageType.JpegImageType);
            //thumb = xpsDoc.AddThumbnail(XpsImageType.JpegImageType);

            //WriteStream(res.GetStream(), resFile);
            //WritePageContent(fp.XmlWriter, res, img.Width, img.Height);
            //res.Commit();

            //WriteStream(thumb.GetStream(), resFile);
            //thumb.Commit();

            //fp.Commit();
            //fd.Commit();
            //fds.Commit();
            //xpsDoc.Close();
            return true;
        }
        //private static void WritePageContent(System.Xml.XmlWriter xmlWriter, XpsResource res, int width, int height)
        //{
        //    xmlWriter.WriteStartElement("FixedPage");
        //    xmlWriter.WriteAttributeString("xmlns", @"http://schemas.microsoft.com/xps/2005/06");
        //    xmlWriter.WriteAttributeString("Width", width.ToString());
        //    xmlWriter.WriteAttributeString("Height", height.ToString());
        //    xmlWriter.WriteAttributeString("xml:lang", "en-US");
        //    xmlWriter.WriteStartElement("Canvas");

        //    if (res is XpsImage)
        //    {
        //        xmlWriter.WriteStartElement("Path");
        //        xmlWriter.WriteAttributeString("Data", "M 20,20 L 770,20 770,770 20,770 z");
        //        xmlWriter.WriteStartElement("Path.Fill");
        //        xmlWriter.WriteStartElement("ImageBrush");
        //        xmlWriter.WriteAttributeString("ImageSource", res.Uri.ToString());
        //        xmlWriter.WriteAttributeString("Viewbox", "0,0,750,750");
        //        xmlWriter.WriteAttributeString("ViewboxUnits", "Absolute");
        //        xmlWriter.WriteAttributeString("Viewport", "20,20,750,750");
        //        xmlWriter.WriteAttributeString("ViewportUnits", "Absolute");
        //        xmlWriter.WriteEndElement();
        //        xmlWriter.WriteEndElement();
        //        xmlWriter.WriteEndElement();
        //    }
        //    xmlWriter.WriteEndElement();
        //    xmlWriter.WriteEndElement();
        //}

        //private static void WriteStream(Stream stream, string resFile)
        //{
        //    using (FileStream sourceStream = new FileStream(resFile, FileMode.Open, FileAccess.Read))
        //    {
        //        byte[] buf = new byte[1024];
        //        int read = 0;
        //        while ((read = sourceStream.Read(buf, 0, buf.Length)) > 0)
        //        {
        //            stream.Write(buf, 0, read);
        //        }
        //    }
        //}

        //public bool Convert(string sourcePath, string targetPath)
        //{
        //    using (FileStream target = new FileStream(targetPath, FileMode.Create))
        //    {
        //        Stream source = new StreamReader(sourcePath).BaseStream;
        //        //Get Bytes from Document stream and write into IO stream
        //        byte[] binaryData = new Byte[source.Length];
        //        long bytesRead = source.Read(binaryData, 0, (int)source.Length);
        //        source.Seek(0, SeekOrigin.Begin);
        //        target.Write(binaryData, 0, binaryData.Length);
        //    }
        //    return true;
        //}
    }
}
