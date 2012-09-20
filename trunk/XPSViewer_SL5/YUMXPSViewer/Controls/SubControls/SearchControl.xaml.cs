using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Controls.Primitives;
using YUMXPSViewer.Controls;
//using YUMXPSViewer.Documents.Text;

namespace YUMXPSViewer.Controls.SubControls
{
    /// <summary>
    /// Represents a control that searches for text within the DocumentViewer
    /// </summary>
	public partial class SearchControl : UserControl
	{
        /// <summary>
        /// Creates a new instance of SearchControl
        /// </summary>
		public SearchControl()
		{
			// Required to initialize variables
			InitializeComponent();            
            Application.Current.Host.Content.FullScreenChanged +=new EventHandler(Content_FullScreenChanged);
            
		}

        private void Content_FullScreenChanged(object sender, EventArgs e)
        {
            this.SearchTextBox.IsReadOnly = Application.Current.Host.Content.IsFullScreen;          
        }

		private void SearchButton_Click(object sender, System.Windows.RoutedEventArgs e)
		{
            //DocumentViewer fixedDocumentViewer = ((Button)sender).DataContext as DocumentViewer;
            //if(fixedDocumentViewer.Document != null)
            //    fixedDocumentViewer.SearchTextAsync(SearchTextBox.Text, TextSearch.SearchModes.None, null);
		}

		private void SearchTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
		{			
            //if (e.Key == Key.Enter)
            //{
            //    DocumentViewer fixedDocumentViewer = ((TextBox)sender).DataContext as DocumentViewer;
            //    if (fixedDocumentViewer.Document != null)                    
            //        fixedDocumentViewer.SearchTextAsync(SearchTextBox.Text, TextSearch.SearchModes.None, null);            	

            //}
		}

        private void SearchTextBox_MouseEnter(object sender, MouseEventArgs e)
        {
            if (Application.Current.Host.Content.IsFullScreen)
            {
                TextSearchToolTipBlock.Text = "For security purposes, Silverlight restricts keyboard access during full-screen mode.";
            }
            else
            {
                TextSearchToolTipBlock.Text = "Search text";
            }
        }			
	}
}