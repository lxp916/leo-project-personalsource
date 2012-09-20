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
using System.Windows.Resources;
using System.Windows.Markup;
using System.Windows.Media.Imaging;
using System.Windows.Browser;
using System.Windows.Navigation;

namespace YUMXPSViewer.Utility.XpsDocument
{
    public class FixedPage:Canvas
    {
        internal static event NavigatedEventHandler NavigateTo;
        #region 附加属性
        /// <summary>
        /// 文件链接
        /// </summary>
        internal static readonly DependencyProperty NavigateUriProperty = DependencyProperty.RegisterAttached("NavigateUri", typeof(string), typeof(FixedPage), new PropertyMetadata(""));
        /// <summary>
        /// 文档照片源属性
        /// </summary>
        internal static readonly DependencyProperty ImageSourceProperty = DependencyProperty.RegisterAttached("ImageSource", typeof(string), typeof(FixedPage), null);
        /// <summary>
        /// Glyphs字体属性
        /// </summary>
        internal static readonly DependencyProperty FontUriProperty = DependencyProperty.RegisterAttached("FontUri", typeof(string), typeof(FixedPage), null);
        private static List<ImageBrush> brushs = new List<ImageBrush>();
        public static void SetImageSource(ImageBrush element, string value)
        {
            element.SetValue(FixedPage.FontUriProperty, value);
            //设置图片源
            if (!XpsDocument.Resource.ImageBrushSource.ContainsKey(value))
            {
                BitmapImage image = new BitmapImage();
                image.SetSource(Application.GetResourceStream(XpsDocument.DocumentStream, ConvertPartName(value)).Stream);
                element.ImageSource = image;
                XpsDocument.Resource.ImageBrushSource.Add(value, image);
            }
            else
            {
                element.ImageSource = XpsDocument.Resource.ImageBrushSource[value];
            }
            element.Transform = null;
            brushs.Add(element);
        }
        public static string GetImageSource(ImageBrush element)
        {
            return (string)element.GetValue(FixedPage.FontUriProperty);
        }
        
        public static void SetFontUri(Glyphs element, string value)
        {
            element.SetValue(FixedPage.FontUriProperty, value);
            //设置字体的字体源
            if (!XpsDocument.Resource.FontUrlSource.ContainsKey(value))
            {
                FontSource fontSource = new FontSource(Application.GetResourceStream(XpsDocument.DocumentStream, ConvertPartName(value)).Stream);
               XpsDocument.Resource.FontUrlSource.Add(value, fontSource);
                element.FontSource = fontSource;
            }
            else
            {
                element.FontSource = XpsDocument.Resource.FontUrlSource[value];
            }
        }
        public static string GetFontUri(Glyphs element)
        {
            return (string)element.GetValue(FixedPage.FontUriProperty);
        }

        public static void SetNavigateUri(FrameworkElement element, string value)
        {
            element.SetValue(FixedPage.NavigateUriProperty, value);
            element.Cursor = Cursors.Hand;
            ToolTipService.SetToolTip(element, new TextBlock { Text =value });
            element.MouseLeftButtonDown += new MouseButtonEventHandler(element_MouseLeftButtonDown);
        }
        static void element_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Path p = sender as Path;
            //p.Dispatcher.BeginInvoke(() =>
            //{
            string uri = GetNavigateUri(p);
            int num = uri.IndexOf('#');
            var pageContent = UriIsExist(uri.Substring(num + 1, uri.Length - num - 1));
            if (pageContent == null)
                HtmlPage.Window.Navigate(new Uri(uri), "new");
            else
            {

                if (NavigateTo != null)
                    NavigateTo(null, new NavigationEventArgs(null, pageContent.Source));
            }
            //}
            //    );
        }
        public static string GetNavigateUri(FrameworkElement element)
        {
            return (string)element.GetValue(FixedPage.NavigateUriProperty);
        }
        #endregion

        private static PageContent UriIsExist(string uri)
        {
            foreach (var f in XpsDocument.FixedDocument)
            {
                foreach (var l in f.LinkTargets)
                {
                    if (l.Name == uri)
                        return f;
                }
            }
            return null;
        }
        public FixedPage()
            : base()
        {
        
        }
        internal void LoadPage(Uri path)
        {
            var newStrem = Application.GetResourceStream(XpsDocument.DocumentStream,path);

            using (System.Xml.XmlReader reader = System.Xml.XmlReader.Create(newStrem.Stream))
            {
                brushs = new List<ImageBrush>();
                XpsReaderSetting setting = new XpsReaderSetting();
                XpsToSilverlightXaml xpsReader = new XpsToSilverlightXaml(setting, reader);
                var el = XamlReader.Load(xpsReader.GetXpsFixedPage()) as Canvas;
                foreach (var b in brushs)
                {
                    b.Transform = new MatrixTransform { Matrix = Matrix.Identity };
                }
                this.Height = el.Height;
                this.Width = el.Width;

                this.Children.Clear();
                this.Children.Add(el);
            }
        }
       protected static Uri ConvertPartName(string partName)
       {
           return new Uri(partName.TrimStart(new char[] { '/' }), UriKind.Relative);
       }
    }
}
