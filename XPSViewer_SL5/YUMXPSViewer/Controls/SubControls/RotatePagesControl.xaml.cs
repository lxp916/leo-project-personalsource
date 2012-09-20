using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using YUMXPSViewer.Controls;
using System.Diagnostics;

namespace YUMXPSViewer.Controls.SubControls
{
    public partial class RotatePagesControl : UserControl
    {
        
        public RotatePagesControl(bool enableCCW = false)
        {
            InitializeComponent();

            if (enableCCW)
                this.RotateCounterClockwise.Visibility = System.Windows.Visibility.Visible;
        }

        private void RotateClockwise_Click(object sender, RoutedEventArgs e)
        {
            //if ((sender as FrameworkElement).DataContext as DocumentViewer != null)
            //    ((sender as FrameworkElement).DataContext as DocumentViewer).RotateClockwise();
        }

        private void RotateCounterClockwise_Click(object sender, RoutedEventArgs e)
        {
            //if ((sender as FrameworkElement).DataContext as DocumentViewer != null)
            //    ((sender as FrameworkElement).DataContext as DocumentViewer).RotateCounterClockwise();
        }
    }
}
