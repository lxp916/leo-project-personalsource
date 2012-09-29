using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Xml;
using System.Text;
using System.Collections.Generic;

namespace MyControl.XpsDocument
{
    internal class XpsToSilverlightXaml
    {
        private XpsReaderSetting Setting;
        private XmlReader reader;
        public XpsToSilverlightXaml(XpsReaderSetting setting, XmlReader reader)
        {
            this.Setting = setting;
            this.reader = reader;
        }
        public string GetXpsFixedPage()
        {
            StringBuilder output = new StringBuilder();
            XmlWriterSettings ws = new XmlWriterSettings();
            ws.Indent = true;
            using (XmlWriter writer = XmlWriter.Create(output, ws))
            {
                reader.ReadToDescendant("FixedPage");
                if (reader.NodeType == XmlNodeType.Attribute)
                {
                    writer.WriteStartAttribute(reader.Prefix, reader.LocalName, reader.NamespaceURI);
                    string name = reader.Name;
                    while (reader.ReadAttributeValue())
                    {
                        if (reader.NodeType == XmlNodeType.EntityReference)
                        {
                            writer.WriteEntityRef(reader.Name);
                        }
                        else
                        {
                            writer.WriteString(reader.Value);
                        }
                    }
                    reader.MoveToAttribute(name);
                    writer.WriteEndAttribute();
                }
                else
                {
                    WriteNode(reader, false, writer);
                }

                writer.WriteEndElement();
            }
            reader.Close();
            return output.ToString();
        }
        private char[] writeNodeBuffer;
        private void ReadFixedDocumentSequence()
        {

        }
        private void WriteNode(XmlReader reader, bool defattr, XmlWriter writer)
        {
            if (reader == null)
            {
                throw new ArgumentNullException("reader");
            }
            bool canReadValueChunk = reader.CanReadValueChunk;
            int num = (reader.NodeType == XmlNodeType.None) ? -1 : reader.Depth;
            do
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        if (reader.LocalName == "FixedPage")
                        {
                            writer.WriteStartElement("", "Canvas", "http://schemas.microsoft.com/winfx/2006/xaml/presentation");
                            reader.MoveToAttribute("Height");
                            reader.MoveToAttribute("Name","LayoutRoot");
                            writer.WriteAttributeString("Height", reader.Value);

                            reader.MoveToAttribute("Width");
                            writer.WriteAttributeString("Width", reader.Value);
                        }
                        else
                        {
                            writer.WriteStartElement(reader.LocalName);
                            WriteAttributes(reader, defattr, writer);
                            if (reader.IsEmptyElement)
                            {
                                writer.WriteEndElement();
                            }
                        }
                        break;
                    case XmlNodeType.Text:
                        int num2;
                        if (!canReadValueChunk)
                        {
                            writer.WriteString(reader.Value);
                            break;
                        }
                        if (writeNodeBuffer == null)
                        {
                            writeNodeBuffer = new char[0x400];
                        }
                        while ((num2 = reader.ReadValueChunk(writeNodeBuffer, 0, 0x400)) > 0)
                        {
                            writer.WriteChars(writeNodeBuffer, 0, num2);
                        }
                        break;
                    case XmlNodeType.CDATA:
                        writer.WriteCData(reader.Value);
                        break;
                    case XmlNodeType.EntityReference:
                        writer.WriteEntityRef(reader.Name);
                        break;
                    case XmlNodeType.ProcessingInstruction:
                    case XmlNodeType.XmlDeclaration:
                        writer.WriteProcessingInstruction(reader.Name, reader.Value);
                        break;
                    case XmlNodeType.Comment:
                        writer.WriteComment(reader.Value);
                        break;
                    case XmlNodeType.DocumentType:
                        writer.WriteDocType(reader.Name, reader.GetAttribute("PUBLIC"), reader.GetAttribute("SYSTEM"), reader.Value);
                        break;
                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        writer.WriteWhitespace(reader.Value);
                        break;
                    case XmlNodeType.EndElement:
                        if (reader.LocalName == "FixedPage")
                        {

                        }
                        else
                            writer.WriteFullEndElement();
                        break;
                }
            }
            while (reader.Read() && ((num < reader.Depth) || ((num == reader.Depth) && (reader.NodeType == XmlNodeType.EndElement))));
        }
        private void WriteAttributes(XmlReader reader, bool defattr, XmlWriter writer)
        {
            if (reader == null)
            {
                throw new ArgumentNullException("reader");
            }
            if ((reader.NodeType == XmlNodeType.Element) || (reader.NodeType == XmlNodeType.XmlDeclaration))
            {
                if (reader.MoveToFirstAttribute())
                {
                    WriteAttributes(reader, defattr, writer);
                    reader.MoveToElement();
                }
            }
            else
            {
                if (reader.NodeType != XmlNodeType.Attribute)
                {
                    throw new XmlException("Xml_InvalidPosition");
                }
                do
                {
                    if (!reader.IsDefault)
                    {
                        if (!Setting.RemoveAttribute.Contains(reader.LocalName))
                        {
                            switch (reader.LocalName)
                            {
                                case "FontUri":
                                    writer.WriteAttributeString("my", "FixedPage.FontUri", "clr-namespace:MyControl.XpsDocument;assembly=MyControl", reader.Value);
                                    break;
                                case "ImageSource":
                                    writer.WriteAttributeString("my", "FixedPage.ImageSource", "clr-namespace:MyControl.XpsDocument;assembly=MyControl", reader.Value);
                                    break;
                                case "FixedPage.NavigateUri":
                                    writer.WriteAttributeString("my", "FixedPage.NavigateUri", "clr-namespace:MyControl.XpsDocument;assembly=MyControl", reader.Value);
                                    break;
                                default:
                                    writer.WriteAttributeString(reader.LocalName, reader.Value);
                                    break;
                            }


                        }
                    }
                }
                while (reader.MoveToNextAttribute());
            }
        }

    }
}
