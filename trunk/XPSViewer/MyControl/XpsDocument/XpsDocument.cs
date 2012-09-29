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
using System.Windows.Resources;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using MyControl.Behavior;
using System.Windows.Printing;
using System.IO;
//using MyControl.RIAServices;

namespace MyControl.XpsDocument
{
    public class XpsDocument :Canvas
    {
        #region DependencyProperty
        public static readonly DependencyProperty TrunAnimaProperty = DependencyProperty.Register("TrunAnima", typeof(TurnAnimaBase), typeof(XpsDocument), new PropertyMetadata(new Turn180()));
        public TurnAnimaBase TrunAnima
        {

            get
            {
                return (TurnAnimaBase)this.GetValue(XpsDocument.TrunAnimaProperty);
            }
            set
            {
                this.SetValue(XpsDocument.TrunAnimaProperty, value);
                System.Windows.Interactivity.Interaction.GetBehaviors(this).Add(value);
            }
        }
        public static readonly DependencyProperty IsEnableAnimaProperty = DependencyProperty.Register("IsEnableAnima", typeof(bool), typeof(XpsDocument), new PropertyMetadata(false));
        public bool IsEnableAnima
        {
            get
            {
                return (bool)this.GetValue(XpsDocument.IsEnableAnimaProperty);
            }
            set
            {
                this.SetValue(XpsDocument.IsEnableAnimaProperty, value);
            }
        }
        public static readonly DependencyProperty IsEnableNavigateProperty = DependencyProperty.Register("IsEnableNavigate", typeof(bool), typeof(XpsDocument), new PropertyMetadata(false));
        /// <summary>
        /// 是否启用导航功能
        /// </summary>
        public bool IsEnableNavigate
        {
            get
            {
                return (bool)this.GetValue(XpsDocument.IsEnableNavigateProperty);
            }
            set
            {
                this.SetValue(XpsDocument.IsEnableNavigateProperty, value);
                FixedPage.NavigateTo -= FixedPage_NavigateTo;
                FixedPage.NavigateTo += new System.Windows.Navigation.NavigatedEventHandler(FixedPage_NavigateTo);
            }
        }

        //public static readonly DependencyProperty DisplayLayoutModeProperty = DependencyProperty.Register("DisplayLayoutMode", typeof(LayoutMode), typeof(XpsDocument), new PropertyMetadata(false));

        private LayoutMode displayLayout = LayoutMode.SinglePage;
        /// <summary>
        /// display single page/ double page/.....
        /// </summary>
        public LayoutMode DisplayLayoutMode
        {
            //get
            //{
            //    return (bool)this.GetValue(XpsDocument.DisplayLayoutModeProperty);
            //}
            //set
            //{
            //    this.SetValue(XpsDocument.DisplayLayoutModeProperty, value);
            //}

            get { return this.displayLayout; }

            set
            {
                this.displayLayout = value;
                if (PageLayoutChanged != null)
                    PageLayoutChanged(this, null);
                this.CurrentPageNum = 1;
                //if (ReadFixedPage(this.currentPageNum))
                //{
                //    this.displayLayout = value;
                //    if (PageLayoutChanged != null)
                //        PageLayoutChanged(this, null);
                //}
            }
        }
        #endregion

        internal List<FixedPage> FixedPages;
        /// <summary>
        /// 翻页的时候触发
        /// </summary>
        public  event EventHandler FixedPageChanged;
        public event EventHandler PageLayoutChanged;
        ///// <summary>
        ///// 流初始化完毕
        ///// </summary>
        //public  event EventHandler DocumentLoaded;

        private int displaynumber = 1;
        public XpsDocument(): base()
        {
            FixedPages = new List<FixedPage>();
            this.TrunAnima = new Turn180();
            displaynumber = 1;
        }

        #region Property

        #region 静态属性
        private static XpsDocumentType xpsDocumentType;
        internal static XpsDocumentType XpsDocumentType
        {
            get { return xpsDocumentType; }
        }
        private static FixedDocumentSequence fixedDocumentSequence = null;
        internal static FixedDocumentSequence FixedDocumentSequence
        {
            get
            {
                if (fixedDocumentSequence == null)
                    fixedDocumentSequence = ReadFixedDocumentSequence();
                return fixedDocumentSequence;
            }
        }
        private static FixedDocument fixedDocument = null;
        internal static FixedDocument FixedDocument
        {
            get
            {
                if (fixedDocument == null)
                    fixedDocument = ReadFixedDocument(XpsDocument.FixedDocumentSequence);
                return fixedDocument;
            }
        }
        private static Resource resource;
        internal static Resource Resource
        {
            get
            {
                if (resource == null)
                    resource = new Resource();
                return resource;
            }
        }

        internal static StreamResourceInfo DocumentStream { get; set; }
        #endregion
        /// <summary>
        /// 文档的页面总数
        /// </summary>
        public int PageCount
        {
            get
            {
                int mode = (int)this.displayLayout == 0 ? XpsDocument.FixedDocument.Count : (int)this.displayLayout;
                int intPage = (int)(XpsDocument.FixedDocument.Count / mode);
                double dblpages = ((double)(XpsDocument.FixedDocument.Count)) / mode;
                return dblpages > intPage ? intPage + 1 : intPage;
            }
        }

        private int currentPageNum = 0;
        /// <summary>
        /// 当前的页数
        /// </summary>
        public int CurrentPageNum
        {
            get { return this.currentPageNum ; }

            set
            {
                if (ReadFixedPage(value))
                {
                    this.currentPageNum = value;
                    if (FixedPageChanged != null)
                        FixedPageChanged(this, null);
                }
            }

        }

        #endregion

        #region Methods

        #region private

        private static FixedDocumentSequence ReadFixedDocumentSequence()
        {
            if (DocumentStream == null)
                throw new ArgumentNullException("流为空，请设置xps文档流");
            //这里判断xps文档的类型
            var fixedDocSeq = Application.GetResourceStream(DocumentStream, new Uri("FixedDocSeq.fdseq", UriKind.Relative));
            if (fixedDocSeq == null)
            {
                xpsDocumentType = MyControl.XpsDocument.XpsDocumentType.Print;
                fixedDocSeq = Application.GetResourceStream(DocumentStream, new Uri("/FixedDocumentSequence.fdseq", UriKind.Relative));
            }
            else
            {
                xpsDocumentType = MyControl.XpsDocument.XpsDocumentType.OfficeSaveAs;
            }
            FixedDocumentSequence fixedDocumentSequence = new FixedDocumentSequence();
            using (XmlReader reader = XmlReader.Create(fixedDocSeq.Stream))
            {
                reader.ReadToDescendant("FixedDocumentSequence");
                fixedDocumentSequence.xmlns = reader.NamespaceURI;
                reader.ReadToDescendant("DocumentReference");
                do
                {
                    reader.MoveToAttribute("Source");
                    fixedDocumentSequence.Add(new DocumentReference { Source = ConvertPartName(reader.Value) });
                } while (reader.ReadToNextSibling("DocumentReference"));
            }
            return fixedDocumentSequence;
        }

        private static FixedDocument ReadFixedDocument(FixedDocumentSequence fixedDocumentSequence)
        {
            if (DocumentStream == null)
                throw new ArgumentNullException("流为空，请设置xps文档流");
            Uri baseUri = fixedDocumentSequence[0].Source;
            var fixedDocumentStream = Application.GetResourceStream(DocumentStream, baseUri);
            FixedDocument fixedDocument = new FixedDocument();
            XElement root = XElement.Load(fixedDocumentStream.Stream);
            string u = baseUri.ToString();
            int num = u.LastIndexOf("/");
            string baseStr = u.Substring(0, num + 1);
            foreach (var pageContentElement in root.Elements())
            {
                PageContent pageContent = new PageContent();
                switch (xpsDocumentType)
                {
                    case MyControl.XpsDocument.XpsDocumentType.OfficeSaveAs:
                        pageContent.Source = new Uri(baseStr + pageContentElement.Attribute("Source").Value, UriKind.Relative);
                        break;
                    case MyControl.XpsDocument.XpsDocumentType.Print:
                        pageContent.Source = ConvertPartName(pageContentElement.Attribute("Source").Value);
                        break;
                }
                fixedDocument.Add(pageContent);
                if (pageContentElement.HasElements)
                {
                    foreach (var linkTargetElement in pageContentElement.Elements().First().Elements())
                    {
                        pageContent.LinkTargets.Add(new LinkTarget { Name = linkTargetElement.Attribute("Name").Value });
                    }
                }
            }
            return fixedDocument;
        }

        private bool ReadFixedPage(int pageNum)
        {
            if (pageNum >= 1 && pageNum <= this.PageCount)
            {
                if (IsEnableAnima)
                {
                    TrunAnima.Start();
                }
                this.Children.Clear();
                this.FixedPages.Clear();
                int mode = (int)this.displayLayout == 0 ? XpsDocument.FixedDocument.Count : (int)this.displayLayout;
                double width= 0;
                double height = 0;
                for (var i = 0; i < mode; i++)
                { 
                    FixedPage page = new FixedPage();
                    int realPage = mode * (pageNum - 1) + i;
                    if (realPage >= XpsDocument.FixedDocument.Count) continue;
                    page.LoadPage(FixedDocument[realPage].Source);
                    width = Math.Max(width, page.Width);
                    height = Math.Max(height, page.Height);
                    
                    this.Children.Add(page);
                    this.FixedPages.Add(page);
                }
                for (var i = 0; i < FixedPages.Count; i++)
                {
                    var item = FixedPages[i];
                    item.Width = width;
                    if (i > 0)
                    {
                        Canvas.SetLeft(item, (width + 1) * (1.0 / mode));

                    }
                    //else Canvas.SetLeft(item, 0);
                    //Canvas.SetTop(item, 0);
                    //(item.RenderTransform as MatrixTransform).Matrix.OffsetX
                    ScaleTransform scaleTransform = new ScaleTransform();
                    scaleTransform.CenterX = width / 2;
                    scaleTransform.CenterY = height / 2;
                    scaleTransform.ScaleX = scaleTransform.ScaleY = 1.0 / mode;
                    item.RenderTransform = scaleTransform;
                    
                }
                this.Width = width;
                this.Height = height;
                return true;
            }
            else
                return false;


        }

        private static Uri ConvertPartName(string partName)
        {
            return new Uri(partName.TrimStart(new char[] { '/' }), UriKind.Relative);
        }
        #endregion

        #region public

        public void SetStream(StreamResourceInfo stream)
        {
            fixedDocument = null;
            fixedDocumentSequence = null;
            currentPageNum = 0;
            resource = new Resource();
            DocumentStream = stream;
        }
        
        public void Print()
        {

            PrintDocument document = new PrintDocument();

            // tell the API what to print
            document.PrintPage += (s, args) =>
            {
                //args.PageVisual = GPrint;
                //Image imagePrint = new Image();
                //imagePrint.Source = img.Source;
                //imagePrint.Height = e.PrintableArea.Height;
                //imagePrint.Width = e.PrintableArea.Width;
                args.PageVisual = this;
                args.HasMorePages = true;
            };

            // call the Print() with a proper name which will be visible in the Print Queue
            document.Print("XPS Document Print Application Demo");
        }

        public void Save()
        {
            SaveFileDialog sf = new SaveFileDialog();
            sf.Filter = "XPS Files (*.xps)|*.xps";
            if (sf.ShowDialog() == true)
            {
                using (Stream fs = sf.OpenFile())
                {
                    Stream stream = DocumentStream.Stream;
                    //Get Bytes from Document stream and write into IO stream
                    byte[] binaryData = new Byte[stream.Length];
                    long bytesRead = stream.Read(binaryData, 0, (int)stream.Length);
                    stream.Seek(0, SeekOrigin.Begin);
                    fs.Write(binaryData, 0, binaryData.Length);
                }
            }
        }
        #endregion

        #endregion

        #region Events
        void FixedPage_NavigateTo(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            for (int i = 0; i < FixedDocument.Count; i++)
            {
                if (FixedDocument[i].Source == e.Uri)
                {
                    ReadFixedPage(i + 1);
                    break;
                }
            }
        }
        #endregion

    }
    /// <summary>
    /// xps文档的类型
    /// xps虚拟打印机生成的文档里面的source是绝对路径
    /// Office文档另存为生成的文档里面的source是相对路径
    /// </summary>
    internal enum XpsDocumentType
    {
        /// <summary>
        /// xps文档是通过xps虚拟打印机生成的
        /// </summary>
        Print,
        /// <summary>
        /// xps文档时通过Office文档另存为生成的
        /// </summary>
        OfficeSaveAs
    }

    public enum LayoutMode
    {
        Thumbnail = 0,
        SinglePage,
        DoublePage
    }
}
