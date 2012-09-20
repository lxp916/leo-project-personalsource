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

namespace YUMXPSViewer.Utility.XpsDocument
{
    internal class Resource
    {
        public ImageBrushSource ImageBrushSource { get; set; }
        public FontUrlSource FontUrlSource { get; set; }
        public Resource()
        {
            ImageBrushSource = new ImageBrushSource();
            FontUrlSource = new FontUrlSource();
        }
    }
}
