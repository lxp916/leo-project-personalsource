﻿#pragma checksum "C:\Users\liaxiaop\Desktop\D\Work\SVN Work Space\GoogleSVN\XPSViewer\Test\Views\XPSViewerPage.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "85FD9EC285791D009B61BF8D75B6769C"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.269
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using MyControl.XpsDocument;
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


namespace Test {
    
    
    public partial class XPSViewerPage : System.Windows.Controls.Page {
        
        internal System.Windows.Controls.Grid LayoutRoot;
        
        internal System.Windows.Controls.Border ToolBorder;
        
        internal System.Windows.Controls.Button OpenLocalFileButton;
        
        internal MyControl.XpsDocument.XpsControl xpsControl;
        
        internal System.Windows.Controls.Button btnCloase;
        
        internal System.Windows.Controls.Grid LowerGrid;
        
        internal System.Windows.Controls.Border DocViewerBorder;
        
        internal MyControl.XpsDocument.XpsDocument xpsDocument;
        
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
            System.Windows.Application.LoadComponent(this, new System.Uri("/Test;component/Views/XPSViewerPage.xaml", System.UriKind.Relative));
            this.LayoutRoot = ((System.Windows.Controls.Grid)(this.FindName("LayoutRoot")));
            this.ToolBorder = ((System.Windows.Controls.Border)(this.FindName("ToolBorder")));
            this.OpenLocalFileButton = ((System.Windows.Controls.Button)(this.FindName("OpenLocalFileButton")));
            this.xpsControl = ((MyControl.XpsDocument.XpsControl)(this.FindName("xpsControl")));
            this.btnCloase = ((System.Windows.Controls.Button)(this.FindName("btnCloase")));
            this.LowerGrid = ((System.Windows.Controls.Grid)(this.FindName("LowerGrid")));
            this.DocViewerBorder = ((System.Windows.Controls.Border)(this.FindName("DocViewerBorder")));
            this.xpsDocument = ((MyControl.XpsDocument.XpsDocument)(this.FindName("xpsDocument")));
        }
    }
}
