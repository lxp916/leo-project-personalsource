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
using System.Collections.Generic;

namespace YUMXPSViewer.Utility.XpsDocument
{
    internal class PageContent
    {
        public Uri Source { get; set; }
        public List<LinkTarget> LinkTargets { get; set; }
        internal PageContent()
        {
            LinkTargets = new List<LinkTarget>();
        }
    }
}
