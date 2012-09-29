﻿#pragma checksum "C:\Users\liaxiaop\Desktop\D\Work\SVN Work Space\GoogleSVN\XPSViewer\MyControl\XpsDocument\XpsControl.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "3866F32F1302D8C729A59925E40EBFA5"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.269
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using MyControl.XpsDocument.SubControls;
using System;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Resources;
using System.Windows.Shapes;
using System.Windows.Threading;


namespace MyControl.XpsDocument {
    
    
    public partial class XpsControl : System.Windows.Controls.StackPanel {
        
        internal System.Windows.Controls.StackPanel ToolStackPanel;
        
        internal MyControl.XpsDocument.SubControls.LayoutControl btnLayout;
        
        internal System.Windows.Controls.Button btnPrevious;
        
        internal System.Windows.Controls.Button btnNext;
        
        internal System.Windows.Controls.TextBox txtCurrentPageNumber;
        
        internal System.Windows.Controls.TextBlock PageNumToolTextBlock;
        
        internal System.Windows.Controls.TextBlock Divider;
        
        internal System.Windows.Controls.TextBlock txtTotalPageCount;
        
        internal System.Windows.Controls.Slider documentScale;
        
        internal System.Windows.Controls.TextBox txtZoom;
        
        internal System.Windows.Controls.TextBlock ZoomToolTextBlock;
        
        internal System.Windows.Controls.TextBox txtSearch;
        
        internal System.Windows.Controls.TextBlock TextSearchToolTipBlock;
        
        internal System.Windows.Controls.Button btnSearch;
        
        internal System.Windows.Controls.ProgressBar progressBar;
        
        internal System.Windows.Controls.Button btnDownload;
        
        internal System.Windows.Controls.Button btnPrint;
        
        internal System.Windows.Controls.Button btnRotateCounterClockwise;
        
        internal System.Windows.Controls.Button btnRotateClockwise;
        
        internal System.Windows.Controls.Button btnFullScreen;
        
        internal System.Windows.Controls.Button btnThumb;
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Windows.Application.LoadComponent(this, new System.Uri("/MyControl;component/XpsDocument/XpsControl.xaml", System.UriKind.Relative));
            this.ToolStackPanel = ((System.Windows.Controls.StackPanel)(this.FindName("ToolStackPanel")));
            this.btnLayout = ((MyControl.XpsDocument.SubControls.LayoutControl)(this.FindName("btnLayout")));
            this.btnPrevious = ((System.Windows.Controls.Button)(this.FindName("btnPrevious")));
            this.btnNext = ((System.Windows.Controls.Button)(this.FindName("btnNext")));
            this.txtCurrentPageNumber = ((System.Windows.Controls.TextBox)(this.FindName("txtCurrentPageNumber")));
            this.PageNumToolTextBlock = ((System.Windows.Controls.TextBlock)(this.FindName("PageNumToolTextBlock")));
            this.Divider = ((System.Windows.Controls.TextBlock)(this.FindName("Divider")));
            this.txtTotalPageCount = ((System.Windows.Controls.TextBlock)(this.FindName("txtTotalPageCount")));
            this.documentScale = ((System.Windows.Controls.Slider)(this.FindName("documentScale")));
            this.txtZoom = ((System.Windows.Controls.TextBox)(this.FindName("txtZoom")));
            this.ZoomToolTextBlock = ((System.Windows.Controls.TextBlock)(this.FindName("ZoomToolTextBlock")));
            this.txtSearch = ((System.Windows.Controls.TextBox)(this.FindName("txtSearch")));
            this.TextSearchToolTipBlock = ((System.Windows.Controls.TextBlock)(this.FindName("TextSearchToolTipBlock")));
            this.btnSearch = ((System.Windows.Controls.Button)(this.FindName("btnSearch")));
            this.progressBar = ((System.Windows.Controls.ProgressBar)(this.FindName("progressBar")));
            this.btnDownload = ((System.Windows.Controls.Button)(this.FindName("btnDownload")));
            this.btnPrint = ((System.Windows.Controls.Button)(this.FindName("btnPrint")));
            this.btnRotateCounterClockwise = ((System.Windows.Controls.Button)(this.FindName("btnRotateCounterClockwise")));
            this.btnRotateClockwise = ((System.Windows.Controls.Button)(this.FindName("btnRotateClockwise")));
            this.btnFullScreen = ((System.Windows.Controls.Button)(this.FindName("btnFullScreen")));
            this.btnThumb = ((System.Windows.Controls.Button)(this.FindName("btnThumb")));
        }
    }
}

