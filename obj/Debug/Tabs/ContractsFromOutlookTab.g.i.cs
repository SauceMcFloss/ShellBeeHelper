﻿#pragma checksum "..\..\..\Tabs\ContractsFromOutlookTab.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "4A4BC675FF961C941FF955974C9CC7EE921E578FC7068AF15F5663675A740E28"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using ShellBeeHelper.Tabs;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace ShellBeeHelper.Tabs {
    
    
    /// <summary>
    /// ContractsFromOutlookTab
    /// </summary>
    public partial class ContractsFromOutlookTab : System.Windows.Controls.UserControl, System.Windows.Markup.IComponentConnector {
        
        
        #line 21 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label EmailAddressLabel;
        
        #line default
        #line hidden
        
        
        #line 23 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox EmailAddressTextBox;
        
        #line default
        #line hidden
        
        
        #line 26 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label SourceFolderLabel;
        
        #line default
        #line hidden
        
        
        #line 28 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox SourceFolderTextBox;
        
        #line default
        #line hidden
        
        
        #line 30 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label DestFolderLabel;
        
        #line default
        #line hidden
        
        
        #line 32 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox DestFolderTextBox;
        
        #line default
        #line hidden
        
        
        #line 34 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button ScanButton;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/ShellBeeHelper;component/tabs/contractsfromoutlooktab.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.EmailAddressLabel = ((System.Windows.Controls.Label)(target));
            return;
            case 2:
            this.EmailAddressTextBox = ((System.Windows.Controls.TextBox)(target));
            
            #line 24 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
            this.EmailAddressTextBox.LostFocus += new System.Windows.RoutedEventHandler(this.EmailAddressTextBox_LostFocus);
            
            #line default
            #line hidden
            return;
            case 3:
            this.SourceFolderLabel = ((System.Windows.Controls.Label)(target));
            return;
            case 4:
            this.SourceFolderTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 5:
            this.DestFolderLabel = ((System.Windows.Controls.Label)(target));
            return;
            case 6:
            this.DestFolderTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 7:
            this.ScanButton = ((System.Windows.Controls.Button)(target));
            
            #line 36 "..\..\..\Tabs\ContractsFromOutlookTab.xaml"
            this.ScanButton.Click += new System.Windows.RoutedEventHandler(this.ScanButton_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

