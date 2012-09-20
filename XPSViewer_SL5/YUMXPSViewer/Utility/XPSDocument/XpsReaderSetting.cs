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
using System.Collections.Generic;

namespace YUMXPSViewer.Utility.XpsDocument
{
    internal class XpsReaderSetting
    {

        public List<string> RemoveAttribute { get; set; }
        public XpsReaderSetting()
        {
            this.RemoveAttribute = new List<string>();

            RemoveAttribute.Add("BidiLevel");
            RemoveAttribute.Add("Viewbox");
            RemoveAttribute.Add("TileMode");
            RemoveAttribute.Add("ViewboxUnits");
            RemoveAttribute.Add("ViewportUnits");
            RemoveAttribute.Add("Viewport");
            RemoveAttribute.Add("xmlns");
            RemoveAttribute.Add("lang");
        }
        public XpsReaderSetting(List<string> removeAttribute)
        {
            this.RemoveAttribute = removeAttribute;
        }
    }
}
