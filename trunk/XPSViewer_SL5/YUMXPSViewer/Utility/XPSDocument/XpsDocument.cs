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


namespace YUMXPSViewer.Utility.XpsDocument
{
    public class XpsDocument :Canvas
    {
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
        public static readonly DependencyProperty IsEnableNavigateProperty = DependencyProperty.Register("IsEnableNavigate", typeof(bool), typeof(XpsDocument),new PropertyMetadata(false));
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
        public XpsDocument():base()
        {
            FixedPages = new List<FixedPage>();
            this.TrunAnima = new Turn180();
           
        }
        internal List<FixedPage> FixedPages;
        /// <summary>
        /// 翻页的时候触发
        /// </summary>
        public  event EventHandler FixedPageChanged;
        ///// <summary>
        ///// 流初始化完毕
        ///// </summary>
        //public  event EventHandler DocumentLoaded;
        /// <summary>
        /// 文档的页面总数
        /// </summary>
        public int PageCount
        {
            get
            {
                return XpsDocument.FixedDocument.Count;
            }
        }
        private int currentPageNum=0;
        /// <summary>
        /// 当前的页数
        /// </summary>
        public int CurrentPageNum { get { return this.currentPageNum; }

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
                    fixedDocument = ReadFixedDocument( XpsDocument.FixedDocumentSequence);
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

        public  void SetStream(StreamResourceInfo stream)
        {
            fixedDocument = null;
            fixedDocumentSequence = null;
            currentPageNum = 0;
            resource = new Resource();
            DocumentStream = stream;
        }
        private static FixedDocumentSequence ReadFixedDocumentSequence()
        {
            if (DocumentStream == null)
                throw new ArgumentNullException("流为空，请设置xps文档流");
            //这里判断xps文档的类型
            var fixedDocSeq = Application.GetResourceStream(DocumentStream, new Uri("FixedDocSeq.fdseq", UriKind.Relative));
            if (fixedDocSeq == null)
            {
                xpsDocumentType = YUMXPSViewer.Utility.XpsDocument.XpsDocumentType.Print;
                fixedDocSeq = Application.GetResourceStream(DocumentStream, new Uri("/FixedDocumentSequence.fdseq", UriKind.Relative));
            }
            else
            {
                xpsDocumentType = YUMXPSViewer.Utility.XpsDocument.XpsDocumentType.OfficeSaveAs;
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
                    case YUMXPSViewer.Utility.XpsDocument.XpsDocumentType.OfficeSaveAs:
                        pageContent.Source = new Uri(baseStr + pageContentElement.Attribute("Source").Value, UriKind.Relative);
                        break;
                    case YUMXPSViewer.Utility.XpsDocument.XpsDocumentType.Print:
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
        void FixedPage_NavigateTo(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            for (int i = 0; i < FixedDocument.Count; i++)
            {
                if (FixedDocument[i].Source == e.Uri)
                {
                    ReadFixedPage(i+1);
                    break;
                }
            }
        }
        private bool  ReadFixedPage(int pageNum)
        {
            if (pageNum >= 1 && pageNum <= FixedDocument.Count)
            {
                if (IsEnableAnima)
                {
                    TrunAnima.Start();
                }
                FixedPage page = new FixedPage();
                page.LoadPage(FixedDocument[pageNum - 1].Source);
                this.Children.Clear();
                this.Width = page.Width;
                this.Height = page.Height;
                this.Children.Add(page);
                 return true;
            }else
                return false;
           
            
        }

        private static Uri ConvertPartName(string partName)
        {
            return new Uri(partName.TrimStart(new char[] { '/' }), UriKind.Relative);
        }

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
}
