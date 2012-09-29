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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Resources;
using System.Windows.Browser;

namespace Test
{
    public partial class XPSViewerPage : Page
    {
        public XPSViewerPage()
        {
            InitializeComponent(); 
        }

        // Executes when the user navigates to this page.
        protected override void OnNavigatedTo(NavigationEventArgs e)
        {

        }

        private void OpenLocalFileButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog opFile = new OpenFileDialog();
                opFile.Multiselect = false;
                opFile.Filter = "XPS Files (*.xps)|*.xps";
                if (opFile.ShowDialog() == true)
                {
                    var newStream = new StreamResourceInfo(opFile.File.OpenRead(), null);
                    xpsDocument.SetStream(newStream);
                    xpsControl.Document = xpsDocument;
                    xpsDocument.CurrentPageNum = 1;
                }
            }
            catch(OutOfMemoryException ex)
            {
                ChildWindow errorWin = new ErrorWindow("The File is too large.",ex.Message);
                errorWin.Show();
            }
        }
        private void btnCloase_Click(object sender, RoutedEventArgs e)
        {
            HtmlWindow html = HtmlPage.Window;
            html.Navigate(new Uri("TestTestPage.aspx#/Welcome", UriKind.Relative));//相对路径
        }
    }
}