﻿#pragma checksum "C:\Users\liaxiaop\Desktop\D\Work\SVN Work Space\GoogleSVN\XPSViewer\MyControl\XpsDocument\SubControls\LayoutControl.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "17A5628F379103FD8A488C6D7EE95DB9"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.269
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

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


namespace MyControl.XpsDocument.SubControls {
    
    
    public partial class LayoutControl : System.Windows.Controls.UserControl {
        
        internal System.Windows.Controls.Grid LayoutRoot;
        
        internal System.Windows.Controls.Button btnLayoutButton;
        
        internal System.Windows.Controls.Image LayoutIcon;
        
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
            System.Windows.Application.LoadComponent(this, new System.Uri("/MyControl;component/XpsDocument/SubControls/LayoutControl.xaml", System.UriKind.Relative));
            this.LayoutRoot = ((System.Windows.Controls.Grid)(this.FindName("LayoutRoot")));
            this.btnLayoutButton = ((System.Windows.Controls.Button)(this.FindName("btnLayoutButton")));
            this.LayoutIcon = ((System.Windows.Controls.Image)(this.FindName("LayoutIcon")));
        }
    }
}

