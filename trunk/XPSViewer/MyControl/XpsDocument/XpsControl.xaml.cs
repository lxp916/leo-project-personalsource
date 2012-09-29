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

namespace MyControl.XpsDocument
{
    //缩略图，查找，双页，doc->xps,加密
    public partial class XpsControl : StackPanel
    {
        public static readonly DependencyProperty DocumentProperty = DependencyProperty.Register("Document", typeof(XpsDocument), typeof(XpsControl), null);
        /// <summary>
        /// change page number
        /// </summary>
        public event EventHandler PageNumberChanged;
        public XpsDocument Document
        {
            get
            {
                return (XpsDocument)this.GetValue(XpsControl.DocumentProperty);
            }
            set
            {
                this.SetValue(XpsControl.DocumentProperty, value);
                value.FixedPageChanged -= Document_FixedPageChanged;
                value.FixedPageChanged += new EventHandler(Document_FixedPageChanged);

                List<int> list = new List<int>();
                for (int i = 1; i <= value.PageCount; i++)
                {
                    list.Add(i);
                }
                //pageList.ItemsSource = list;
                //pageList.SelectedIndex = 0;
                value.FixedPageChanged += new EventHandler(value_FixedPageChanged);
                documentScale.ValueChanged -= documentScale_ValueChanged;
                documentScale.ValueChanged += new RoutedPropertyChangedEventHandler<double>(documentScale_ValueChanged);
                this.progressBar.Maximum = value.PageCount;
            }
        }

        public XpsControl()
        {
            InitializeComponent();
            this.btnLayout.PageLayoutChanged += new SubControls.LayoutControl.EventHandler(btnLayout_PageLayoutChanged);
        }

        #region Events
        void value_FixedPageChanged(object sender, EventArgs e)
        {

        }

        void Document_FixedPageChanged(object sender, EventArgs e)
        {
            this.txtCurrentPageNumber.Text = Document.CurrentPageNum.ToString();
            txtTotalPageCount.Text = Document.PageCount.ToString();
            this.progressBar.Value = Document.CurrentPageNum;
        }

        #region Page navigation

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            if (Document != null)
            {
                int origNum = Document.CurrentPageNum;

                Document.CurrentPageNum = Document.CurrentPageNum - 1;

                if (origNum != Document.CurrentPageNum && PageNumberChanged != null)
                    PageNumberChanged(this, new EventArgs());
            }
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            if (Document != null)
            {
                int origNum = Document.CurrentPageNum;

                Document.CurrentPageNum = Document.CurrentPageNum + 1;

                if (origNum != Document.CurrentPageNum && PageNumberChanged != null)
                    PageNumberChanged(this, new EventArgs());
            }
        }

        private void txtCurrentPageNumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (!((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9)) || e.Key == Key.Enter)
            {
                e.Handled = true;
                if (!string.IsNullOrEmpty(txtCurrentPageNumber.Text) && e.Key == Key.Enter
                    && (int.Parse(txtCurrentPageNumber.Text) <= Document.PageCount && int.Parse(txtCurrentPageNumber.Text) > 0))
                {
                    if (Document != null)
                    {
                        int origNum = Document.CurrentPageNum;

                        Document.CurrentPageNum = int.Parse(txtCurrentPageNumber.Text);

                        if (origNum != Document.CurrentPageNum && PageNumberChanged != null)
                            PageNumberChanged(this, new EventArgs());
                    }
                }
            }
            else e.Handled = false;
        }

        private void txtCurrentPageNumber_MouseEnter(object sender, MouseEventArgs e)
        {
            if (Application.Current.Host.Content.IsFullScreen)
            {
                PageNumToolTextBlock.Text = "For security purposes, Silverlight restricts keyboard access during full-screen mode.";
            }
            else
            {
                PageNumToolTextBlock.Text = "Current page number";
            }
        }

        private void txtCurrentPageNumber_LostFocus(object sender, RoutedEventArgs e)
        {
            //((TextBox)sender).GetBindingExpression(TextBox.TextProperty).UpdateSource();
        }
        #endregion

        #region Page Zoom
        private void txtZoom_KeyDown(object sender, KeyEventArgs e)
        {
            if (!((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9)) || e.Key == Key.Enter)
            {
                e.Handled = true;
                if (!string.IsNullOrEmpty(txtZoom.Text) && e.Key == Key.Enter)
                {
                    double value = int.Parse(txtZoom.Text) / 100;
                    if (Document != null)
                    {
                        documentScale.Value = value > documentScale.Maximum ? documentScale.Maximum : value;
                        documentScale_ValueChanged(this, null);
                    }
                }
            }
            else e.Handled = false;
        }

        private void txtZoom_MouseEnter(object sender, MouseEventArgs e)
        {
            if (Application.Current.Host.Content.IsFullScreen)
            {
                PageNumToolTextBlock.Text = "For security purposes, Silverlight restricts keyboard access during full-screen mode.";
            }
            else
            {
                PageNumToolTextBlock.Text = "Current page number";
            }
        }


        private void documentScale_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (Document != null)
            {
                RotateTransform transform = Document.RenderTransform as RotateTransform;
                ScaleTransform scaleTransform = new ScaleTransform();

                double oldWidth = Document.Width;
                double oldHeight = Document.Height;
                scaleTransform.CenterX = oldWidth / 2;
                scaleTransform.CenterY = oldHeight / 2;

                if (transform != null)
                {
                    scaleTransform.CenterX = transform.CenterX;
                    scaleTransform.CenterY = transform.CenterY;
                }
                scaleTransform.ScaleX = scaleTransform.ScaleY = documentScale.Value;
                Document.RenderTransform = scaleTransform;
                this.txtZoom.Text = (documentScale.Value * 100).ToString();
            }
        }

        private void btnFullScreen_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Host.Content.IsFullScreen = !Application.Current.Host.Content.IsFullScreen;
        }
        #endregion

        #region Search
        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.Key == Key.Enter)
            {
                if (Document != null)
                {
                }
                //DocumentViewer fixedDocumentViewer = ((TextBox)sender).DataContext as DocumentViewer;
                //if (fixedDocumentViewer.Document != null)
                //    fixedDocumentViewer.SearchTextAsync(SearchTextBox.Text, TextSearch.SearchModes.None, null);

            }
        }

        private void txtSearch_MouseEnter(object sender, MouseEventArgs e)
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

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //DocumentViewer fixedDocumentViewer = ((Button)sender).DataContext as DocumentViewer;
            //if(fixedDocumentViewer.Document != null)
            //    fixedDocumentViewer.SearchTextAsync(SearchTextBox.Text, TextSearch.SearchModes.None, null);
        }

        #endregion

        #region Tool bar operation
        private void btnDownload_Click(object sender, RoutedEventArgs e)
        {
            if (Document != null)
            {
                Document.Save();
            }
        }

        private void btnThumb_Click(object sender, RoutedEventArgs e)
        {
            if (Document != null)
            {
                Document.DisplayLayoutMode = Document.DisplayLayoutMode == LayoutMode.Thumbnail ? this.btnLayout.Value : LayoutMode.Thumbnail;
            }
        }

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            if (Document != null)
            {
                Document.Print();
            }
        }

        private void btnRotateClockwise_Click(object sender, RoutedEventArgs e)
        {
            if (Document != null)
            {
                RotateTransform transform = Document.RenderTransform as RotateTransform;
                if (transform == null) transform = new RotateTransform();
                transform.Angle = transform.Angle >= 360 ? 0 : transform.Angle;
                double oldWidth = Document.Width;
                double oldHeight = Document.Height;
                double cavansLeft = Canvas.GetLeft(Document);
                double cavansTop = Canvas.GetTop(Document);
                transform.CenterX = oldWidth / 2;
                transform.CenterY = oldHeight / 2;
                Document.Width = oldHeight;
                Document.Height = oldWidth;
                Canvas.SetLeft(Document, cavansTop);
                Canvas.SetTop(Document, cavansLeft);

                transform.Angle += 90.0;
                Document.RenderTransform = transform;
            }
        }

        private void btnRotateCounterClockwise_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnLayout_PageLayoutChanged(object sender, RoutedEventArgs e)
        {
            if (Document == null)
                this.btnLayout.Value = LayoutMode.SinglePage;
            else
                Document.DisplayLayoutMode = this.btnLayout.Value;
        }


        //private void pageList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    if (Document != null)
        //    {
        //        Document.ReadFixedPage(int.Parse(pageList.SelectedItem.ToString()));
        //    }
        //}
        #endregion

        #endregion

        #region Methods

        #endregion
    }
}
