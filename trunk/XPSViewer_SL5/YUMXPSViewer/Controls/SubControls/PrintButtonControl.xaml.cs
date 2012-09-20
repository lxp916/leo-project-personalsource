using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using YUMXPSViewer.Controls;
using YUMXPSViewer.Controls;


namespace YUMXPSViewer.Controls.SubControls
{
	/// <summary>
	/// Represents a button control that brings up the printing prompt on click
	/// </summary>
	public partial class PrintButtonControl : UserControl
	{
        /// <summary>
        /// Creates a new instance of PrintButtonControl
        /// </summary>
		public PrintButtonControl()
		{
			// Required to initialize variables
			InitializeComponent();
		}
		
		private void PrintButton_Click(object sender, RoutedEventArgs e)
		{
            //DocumentViewer fixedDocumentViewer = ((Button)sender).DataContext as DocumentViewer;
            //fixedDocumentViewer.Document.Print();
		}
	}
}