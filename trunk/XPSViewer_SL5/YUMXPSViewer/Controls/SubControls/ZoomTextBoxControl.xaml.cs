using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using YUMXPSViewer.Utility;
using YUMXPSViewer.Controls;
namespace YUMXPSViewer.Controls.SubControls
{
    /// <summary>
    /// Represents a textbox control that indicates and modifies the current zoom level
    /// </summary>
	public partial class ZoomTextBoxControl : UserControl
	{
        /// <summary>
        /// Creates a new instance of ZoomTextBoxControl
        /// </summary>
		public ZoomTextBoxControl()
		{			
			InitializeComponent();
            Application.Current.Host.Content.FullScreenChanged += new EventHandler(Content_FullScreenChanged);
		}
        
        private void Content_FullScreenChanged(object sender, EventArgs e)
        {
            this.ZoomTextBox.IsReadOnly = Application.Current.Host.Content.IsFullScreen;
        }

        private void ZoomTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //DocumentViewer viewer = this.DataContext as DocumentViewer;                          
                //((TextBox)sender).GetBindingExpression(TextBox.TextProperty).UpdateSource();
            }
        }

        private void ZoomTextBox_MouseEnter(object sender, MouseEventArgs e)
        {
            if (Application.Current.Host.Content.IsFullScreen)
            {
                ZoomToolTextBlock.Text = "For security purposes, Silverlight restricts keyboard access during full-screen mode.";
            }
            else
            {
                ZoomToolTextBlock.Text = "Zoom";
            }
        }


	}
}