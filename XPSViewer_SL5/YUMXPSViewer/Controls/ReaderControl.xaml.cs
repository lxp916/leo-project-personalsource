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
using System.IO;
using System.Diagnostics;
using System.ComponentModel;
using System.Windows.Data;
using System.Windows.Resources;

using YUMXPSViewer.Controls;
using YUMXPSViewer.Utility;
using YUMXPSViewer.Utility.XpsDocument;
using YUMXPSViewer.Controls.SubControls;



namespace YUMXPSViewer.Controls
{
    /// <summary>
    /// Represents a control that provides tools for user interaction with the DocumentViewer.
    /// </summary>
    public partial class ReaderControl : UserControl, INotifyPropertyChanged
    {
        //private IPartRetriever myRetriever = null;
        //private bool _toolsCreated;
        //private bool _documentLoaded;
        //private GridLength _sideWindowWidth = new GridLength(200, GridUnitType.Pixel);
        //private const double MIN_SIDE_WINDOW_WIDTH = 110; // minimum size of the TabControl before the TabItems start to wrap vertically

        //#region Toolbar & Sidewindow Creation

        /// <summary>
        /// Determines the creation of the OpenLocalFileControl upon tool bar generation
        /// </summary>       
        public bool EnableOpenLocalFileControl { get; set; }

        /// <summary>
        /// Determines the creation of the PageNumberControl upon tool bar generation
        /// </summary>
        public bool EnablePageNumberControl { get; set; }

        /// <summary>
        /// Determines the creation of the PageNavigationControl upon tool bar generation
        /// </summary>
        public bool EnablePageNavigationControl { get; set; }

        /// <summary>
        /// Determines the creation of the ZoomTextBoxControl; upon tool bar generation
        /// </summary>
        public bool EnableZoomTextBoxControl { get; set; }

        /// <summary>
        /// Determines the creation of the ZoomSliderControl upon tool bar generation
        /// </summary>
        public bool EnableZoomSliderControl { get; set; }

        /// <summary>
        /// Determines the creation of the FitModeControl upon tool bar generation
        /// </summary>
        public bool EnableFitModeControl { get; set; }

        /// <summary>
        /// Determines the creation of the ToolModeControl upon tool bar generation
        /// </summary>
        public bool EnableToolModeControl { get; set; }

        /// <summary>
        /// Determines the creation of the SearchControl upon tool bar generation
        /// </summary>
        public bool EnableSearchControl { get; set; }

        /// <summary>
        /// Determines the creation of the FullScreenControl upon tool bar generation
        /// </summary>
        public bool EnableFullScreenControl { get; set; }

        /// <summary>
        /// Determines the creation of the PrintControl upon tool bar generation
        /// </summary>
        public bool EnablePrintControl { get; set; }

        /// <summary>
        /// Determines the creation of the OutlineToggleButton upon tool bar generation
        /// </summary>
        public bool EnableOutlineToggleControl { get; set; }

        ///// <summary>
        ///// Determines the creation of the ThumbnailListControl upon side windowr generation
        ///// </summary>
        //public bool EnableThumbnailListControl { get; set; }

        ///// <summary>
        ///// Determines the creation of the OutlineTreeControl upon side window generation
        ///// </summary>
        //public bool EnableOutlineTreeControl { get; set; }

        ///// <summary>
        ///// Determines the creation of the EnableFullTextSearchControl upon side window generation
        ///// </summary>
        //public bool EnableFullTextSearchControl { get; set; }

        /// <summary>
        /// Determines the creation of the LayoutControl upon tool bar generation
        /// </summary>
        public bool EnableLayoutControl { get; set; }

        /// <summary>
        /// Determines the creation of the RotateControl upon tool bar generation
        /// </summary>
        public bool EnableRotateControl { get; set; }

        /// <summary>
        /// Notification event that is raised when a bound property changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        ///// <summary>
        ///// Used for opening the search side panel if node is opened from a search result.
        ///// </summary>
        //public String InitialSearchTerm { get; set; }

        /// <summary>
        /// Layout mode enumerations
        /// </summary>
        public enum LayoutModes
        {
            Continuous = 0,
            FacingContinous = 1,
            FacingCoverContinuous = 2,
            SinglePage = 3,
            Facing = 4,
            FacingCover = 5
        }

        private LayoutModes _layoutMode;
        /// <summary>
        /// Determines the page presentation mode of the document
        /// </summary>
        public LayoutModes LayoutMode
        {
            get { return _layoutMode; }
            set
            {
                if (_layoutMode != value)
                {
                    _layoutMode = value;
                    this.SetLayout(_layoutMode);
                    if (PropertyChanged != null)
                        this.PropertyChanged(this, new PropertyChangedEventArgs("LayoutMode"));
                }
            }
        }

        ///// <summary>
        ///// Page View mode enumerations
        ///// </summary>
        //public enum PageViewModes
        //{
        //    /// <summary>
        //    /// Page is zoomed. The zoom ratio is specified using <see cref="DocumentViewer.Zoom"/> property.
        //    /// </summary>
        //    Zoom = 0,

        //    /// <summary>
        //    /// Page zoom is automatically adjusted so that page width fits into available space.
        //    /// </summary>
        //    FitWidth = 1,

        //    /// <summary>
        //    /// Page zoom is automatically adjusted so that page height fits into available space.
        //    /// </summary>
        //    FitHeight = 2,

        //    /// <summary>
        //    /// Page zoom is automatically adjusted so that entire page fits into available space.
        //    /// </summary>
        //    FitPage = 3
        //}


        //private void UpdatePageViewMode()
        //{
        //    if (this.FixedDocViewer.FitModeWidth == DocumentViewer.FitModes.Panel && this.FixedDocViewer.FitModeHeight == DocumentViewer.FitModes.None)
        //        PageViewMode = PageViewModes.FitWidth;
        //    else if (this.FixedDocViewer.FitModeWidth == DocumentViewer.FitModes.Panel && this.FixedDocViewer.FitModeHeight == DocumentViewer.FitModes.Page)
        //        PageViewMode = PageViewModes.FitPage;
        //    else if (this.FixedDocViewer.FitModeWidth == DocumentViewer.FitModes.None && this.FixedDocViewer.FitModeHeight == DocumentViewer.FitModes.Page)
        //        PageViewMode = PageViewModes.FitHeight;
        //    else
        //        PageViewMode = PageViewModes.Zoom;

        //}

        //private PageViewModes _PageViewMode;

        ///// <summary>
        ///// Gets or sets the ReaderControl's PageViewMode
        ///// </summary>
        //public PageViewModes PageViewMode
        //{
        //    get
        //    {
        //        return _PageViewMode;
        //    }
        //    set
        //    {
        //        if (_PageViewMode != value)
        //        {
        //            _PageViewMode = value;

        //            if (PropertyChanged != null)
        //                this.PropertyChanged(this, new PropertyChangedEventArgs("PageViewMode"));

        //            if (_PageViewMode == PageViewModes.FitWidth)
        //                FixedDocViewer.SetFitMode(DocumentViewer.FitModes.Panel, DocumentViewer.FitModes.None);
        //            else if (_PageViewMode == PageViewModes.FitHeight)
        //                FixedDocViewer.SetFitMode(DocumentViewer.FitModes.None, DocumentViewer.FitModes.Page);
        //            else if (_PageViewMode == PageViewModes.FitPage)
        //                FixedDocViewer.SetFitMode(DocumentViewer.FitModes.Panel, DocumentViewer.FitModes.Page);
        //            else if (_PageViewMode == PageViewModes.Zoom)
        //                FixedDocViewer.SetFitMode(DocumentViewer.FitModes.None, DocumentViewer.FitModes.None);
        //        }
        //    }
        //}

        //private void SetToolsAndWindowDefault()
        //{
        //    EnableOutlineToggleControl = true;
        //    EnablePageNumberControl = false;
        //    EnablePageNavigationControl = true;

        //    EnableZoomSliderControl = true;
        //    EnableZoomTextBoxControl = true;
        //    EnableFitModeControl = true;

        //    EnableToolModeControl = true;
        //    EnableFullScreenControl = true;
        //    EnablePrintControl = true;
        //    EnableOpenLocalFileControl = true;
        //    EnableSearchControl = true;

        //    EnableThumbnailListControl = true;
        //    EnableOutlineTreeControl = true;
        //    EnableFullTextSearchControl = true;

        //    EnableLayoutControl = true;
        //    EnableRotateControl = true;
        //}

        //#endregion

        //#region Dependency Properties
        ///// <summary>
        ///// Dependency Property for ShowToolbar
        ///// </summary>
        //public static readonly DependencyProperty ShowToolbarProperty
        //= DependencyProperty.Register("ShowToolbarProperty", typeof(bool),
        //typeof(ReaderControl), new PropertyMetadata(true, new PropertyChangedCallback(OnShowToolbarPropertyChanged)));

        ///// <summary>
        ///// Displays or Hides the tool bar on the side window
        ///// </summary>
        //public bool ShowToolbar
        //{
        //    get { return (bool)GetValue(ShowToolbarProperty); }
        //    set { SetValue(ShowToolbarProperty, value); }
        //}

        //private static void OnShowToolbarPropertyChanged(Object sender, DependencyPropertyChangedEventArgs e)
        //{
        //    ReaderControl source = (ReaderControl)sender;

        //    if ((bool)e.NewValue)
        //    {
        //        source.DocumentToolbar.Visibility = Visibility.Visible;
        //    }
        //    else
        //    {
        //        source.DocumentToolbar.Visibility = Visibility.Collapsed;
        //    }
        //}

        /// <summary>
        /// Dependency Property for ShowSideWindow
        /// </summary>
        public static readonly DependencyProperty ShowSideWindowProperty = DependencyProperty.Register("ShowSideWindowProperty", typeof(bool),
        typeof(ReaderControl), new PropertyMetadata(false, new PropertyChangedCallback(OnShowSideWindowPropertyChanged)));

        /// <summary>
        /// Displays or hides the side window.
        /// </summary>
        public bool ShowSideWindow
        {
            get { return (bool)GetValue(ShowSideWindowProperty); }
            set { SetValue(ShowSideWindowProperty, value); }
        }

        private static void OnShowSideWindowPropertyChanged(Object sender, DependencyPropertyChangedEventArgs e)
        {
            ReaderControl source = (ReaderControl)sender;
            source.SetSideWindowVisible((bool)e.NewValue);

        }

        ///// <summary>
        ///// Dependency Property for InitialDocumentUrl
        ///// </summary>
        //public static readonly DependencyProperty InitialDocumentUrlProperty
        //= DependencyProperty.Register("InitialDocumentUrlProperty", typeof(string),
        //typeof(ReaderControl), new PropertyMetadata(null, new PropertyChangedCallback(OnInitialDocumentUrlPropertyChanged)));

        ///// <summary>
        ///// Uri of the document to be loaded initially.
        ///// </summary>
        //public string InitialDocumentUrl
        //{
        //    get { return (string)GetValue(InitialDocumentUrlProperty); }
        //    set { SetValue(InitialDocumentUrlProperty, value); }
        //}

        //private static void OnInitialDocumentUrlPropertyChanged(Object sender, DependencyPropertyChangedEventArgs e)
        //{
        //    if (e.NewValue != null)
        //    {
        //        ReaderControl source = (ReaderControl)sender;
        //        string url = e.NewValue as string;
        //        //source.LoadDocument(url);
        //    }
        //}
        //#endregion


        /// <summary>
        /// Creates a new ReaderControl
        /// </summary>
        public ReaderControl()
        {
            InitializeComponent();

            //_toolsCreated = false;
            //_documentLoaded = false;

            //SetToolsAndWindowDefault();
            //this.FixedDocViewer.PropertyChanged += new PropertyChangedEventHandler(FixedDocViewer_PropertyChanged);

            //this.PageViewMode = PageViewModes.FitWidth;

            //MenuItem dockMenuItem = new MenuItem();
            //dockMenuItem.Header = "Dock on bottom";
            //dockMenuItem.Click += new RoutedEventHandler(dockMenuItem_Click);
            //this.DocumentToolbar.ContextMenu.Items.Add(dockMenuItem);
        }

        //void dockMenuItem_Click(object sender, RoutedEventArgs e)
        //{
        //    if (DocumentToolbar.VerticalAlignment == System.Windows.VerticalAlignment.Top)
        //    {
        //        (sender as MenuItem).Header = "Dock on top";
        //        DocumentToolbar.VerticalAlignment = System.Windows.VerticalAlignment.Bottom;
        //    }
        //    else if (DocumentToolbar.VerticalAlignment == System.Windows.VerticalAlignment.Bottom)
        //    {
        //        (sender as MenuItem).Header = "Dock on botton";
        //        DocumentToolbar.VerticalAlignment = System.Windows.VerticalAlignment.Top;
        //    }
        //}

        //private void FixedDocViewer_PropertyChanged(object sender, PropertyChangedEventArgs e)
        //{
        //    if (e.PropertyName.Equals("FitModeWidth") || e.PropertyName.Equals("FitModeHeight"))
        //    {
        //        UpdatePageViewMode();
        //    }
        //}


        //private void OnLoadAsyncCallback(Exception error)
        //{

        //    if (error != null)
        //    {
        //        String errorString = "Error loading document: ";
        //        WebException webException = error as WebException;
        //        if (webException != null)
        //        {
        //            foreach (string header in webException.Response.Headers)
        //            {
        //                Debug.WriteLine(header);
        //            }
        //        }
        //        if (error.Message != String.Empty)
        //            errorString += error.Message;
        //        if (error.InnerException != null && error.InnerException.Message != String.Empty && error.InnerException.Message != error.Message)
        //            errorString += " " + error.InnerException.Message;

        //        if (error.Message == String.Empty && (error.InnerException == null || error.InnerException.Message == String.Empty))
        //            errorString += "An unknown error has occured.";

        //        MessageBox.Show(errorString);
        //    }
        //    else
        //    {
        //        //document is loaded without error
        //        if (!string.IsNullOrWhiteSpace(this.InitialSearchTerm) && this.EnableSearchControl)
        //        {
        //            this.ShowSideWindow = true;
        //            this.DocumentSideWindow.SelectSearchTabItem(this.InitialSearchTerm);
        //        }
        //    }
        //}

        ///// <summary>
        ///// Loads a remote document through url
        ///// </summary>
        ///// <param name="path">url path in string</param>
        //public void LoadDocument(String path)
        //{
        //    if (System.ComponentModel.DesignerProperties.IsInDesignTool)
        //        return;

        //    if (path == null || path == String.Empty)
        //        return;

        //    if (myRetriever != null)
        //        myRetriever.CancelAllRequests();

        //    Uri uri = new Uri(path, UriKind.RelativeOrAbsolute);

        //    myRetriever = new HttpPartRetriever(uri);
        //    FixedDocViewer.LoadAsync(myRetriever, OnLoadAsyncCallback);
        //    FixedDocViewer.Document.BackgroundLoading = true;
        //    FixedDocViewer.ThumbnailOnScroll = true;
        //}

        ///// <summary>
        ///// Loads a remote document through url
        ///// </summary>
        ///// <param name="uri">url path in Uri</param>
        //public void LoadDocument(Uri uri)
        //{
        //    if (System.ComponentModel.DesignerProperties.IsInDesignTool)
        //        return;

        //    if (uri == null)
        //        return;

        //    myRetriever = new HttpPartRetriever(uri);
        //    FixedDocViewer.LoadAsync(myRetriever, OnLoadAsyncCallback);
        //    FixedDocViewer.Document.BackgroundLoading = true;
        //    FixedDocViewer.ThumbnailOnScroll = true;
        //}

        /// <summary>
        /// Loads a local document from client file system
        /// </summary>
        public void LoadLocalDocument()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = false;
            dlg.Filter = "XPS Files (*.xps)|*.xps";

            // open dialog
            if ((bool)dlg.ShowDialog())
            {
                var newStream = new StreamResourceInfo(dlg.File.OpenRead(), null);

                //if (myRetriever as LocalPartRetriever != null)
                //    ((LocalPartRetriever)myRetriever).Dispose();
                //FileStream fileStream = dlg.File.OpenRead();
                //myRetriever = new LocalPartRetriever(fileStream);
                //FixedDocViewer.LoadAsync(myRetriever, OnLoadAsyncCallback);
                //FixedDocViewer.Document.BackgroundLoading = false;
                //FixedDocViewer.ThumbnailOnScroll = false;
            }

        }


        //public void AddCustomTabItem(TabItem item)
        //{
        //    this.DocumentSideWindow.SideTabControl.Items.Add(item);
        //}
        //public void AddCustomToolItem(FrameworkElement item)
        //{
        //    this.DocumentToolbar.ToolStackPanel.Children.Add(item);
        //}

        //private void UserControl_Loaded(object sender, RoutedEventArgs e)
        //{

        //    if (!_toolsCreated)
        //    {
        //        // Generate toolbar
        //        DocumentToolbar.CreateToolbar(this);
        //        if (this.EnableOutlineTreeControl == true || this.EnableThumbnailListControl == true ||
        //            this.EnableSearchControl == true)
        //        {
        //            this.DocumentSideWindow.CreateSideWindow(this);
        //        }
        //        _toolsCreated = true;
        //        SetSideWindowVisible(this.ShowSideWindow);
        //    }
        //    if (!_documentLoaded && InitialDocumentUrl != null)
        //    {
        //        //load initial document uri specified by xaml
        //        this.LoadDocument(InitialDocumentUrl);
        //        _documentLoaded = true;
        //    }

        //}

        // Toggles the visilbility of the side window.
        // Called when the Dependency Property ShowSideWindow is changed.
        private void SetSideWindowVisible(bool visible)
        {
            //if (visible)
            //{
            //    DocumentSideWindow.Visibility = Visibility.Visible;
            //    SideWindowSplitter.Visibility = Visibility.Visible;
            //    LowerGrid.ColumnDefinitions[0].Width = _sideWindowWidth;
            //    this.DocViewerBorder.Margin = new Thickness(5, 0, 0, 0);
            //}
            //else
            //{
            //    DocumentSideWindow.Visibility = Visibility.Collapsed;
            //    SideWindowSplitter.Visibility = Visibility.Collapsed;

            //    if (LowerGrid.ColumnDefinitions[0].ActualWidth < MIN_SIDE_WINDOW_WIDTH)
            //        _sideWindowWidth = new GridLength(MIN_SIDE_WINDOW_WIDTH, GridUnitType.Pixel);
            //    else
            //        _sideWindowWidth = LowerGrid.ColumnDefinitions[0].Width;

            //    LowerGrid.ColumnDefinitions[0].Width = new GridLength(0, GridUnitType.Pixel);
            //    this.DocViewerBorder.Margin = new Thickness(0, 0, 0, 0);
            //}
        }

        //private void DocumentSideWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        //{
        //    if (e.NewSize.Width < MIN_SIDE_WINDOW_WIDTH)
        //    {
        //        this.ShowSideWindow = false;
        //    }

        //    if (this.DocumentSideWindow.ActualWidth + 8 > this.LayoutRoot.ActualWidth)
        //    {
        //        LowerGrid.ColumnDefinitions[0].Width = new GridLength(this.LayoutRoot.ActualWidth - 8, GridUnitType.Pixel);
        //    }


        //}

        /// <summary>
        ///  Sets the page layout mode of the current document viewer
        /// </summary>
        /// <param name="mode">the LayoutModes to change to</param>
        private void SetLayout(LayoutModes mode)
        {
            //Debug.WriteLine("Entering SetLayout " + DateTime.Now);
            //if (mode == LayoutModes.Continuous)
            //    FixedDocViewer.Template = (ControlTemplate)this.Resources["VerticalLayoutTemplate"];
            //else if (mode == LayoutModes.FacingContinous)
            //    FixedDocViewer.Template = (ControlTemplate)this.Resources["FacingLayoutTemplate"];
            //else if (mode == LayoutModes.FacingCoverContinuous)
            //    FixedDocViewer.Template = (ControlTemplate)this.Resources["FacingCoverContinousLayoutTemplate"];
            //else if (mode == LayoutModes.SinglePage)
            //    FixedDocViewer.Template = (ControlTemplate)this.Resources["VerticalLayoutTemplate"];
            //else if (mode == LayoutModes.FacingCover)
            //    FixedDocViewer.Template = (ControlTemplate)this.Resources["FacingCoverContinousLayoutTemplate"];
            //else if (mode == LayoutModes.Facing)
            //    FixedDocViewer.Template = (ControlTemplate)this.Resources["FacingLayoutTemplate"];

            //if (mode == LayoutModes.SinglePage)
            //{
            //    FixedDocViewer.DisplayMode = DocumentViewer.DisplayModes.SinglePage;
            //}
            //else if (mode == LayoutModes.Facing)
            //{
            //    FixedDocViewer.DisplayMode = DocumentViewer.DisplayModes.DualPageFacing;
            //}
            //else if (mode == LayoutModes.FacingCover)
            //{
            //    FixedDocViewer.DisplayMode = DocumentViewer.DisplayModes.DualPageCoverFacing;
            //}
            //else
            //{
            //    FixedDocViewer.DisplayMode = DocumentViewer.DisplayModes.AllPages;
            //}

            //FixedDocViewer.RefreshTemplate();


        }

        private void DocumentToolbar_IsPinnedChanged(object sender, RoutedPropertyChangedEventArgs<bool> e)
        {
            if (e.NewValue)
            {
                //pin
                Grid.SetRow(DocumentToolbar, 0);
            }
            else
            {
                //unpin
                Grid.SetRow(DocumentToolbar, 1);
            }
        }
    }
}
