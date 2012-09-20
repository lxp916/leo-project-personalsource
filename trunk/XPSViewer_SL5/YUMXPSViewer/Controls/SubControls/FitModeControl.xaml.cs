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

namespace YUMXPSViewer.Controls.SubControls
{
	/// <summary>
	/// Represents a button control that brings up the printing prompt on click
	/// </summary>
	public partial class FitModeControl : UserControl
	{
        /// <summary>
        /// Creates a new instance of FitModeControl
        /// </summary>
        public FitModeControl(ReaderControl readerContorl)
		{
			// Required to initialize variables
			InitializeComponent();
            this.DataContext = readerContorl;
		}
	}
}