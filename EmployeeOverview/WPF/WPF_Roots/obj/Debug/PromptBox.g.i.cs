﻿#pragma checksum "..\..\PromptBox.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "7FBFCDDA77D0FE6D308474F50DDDEFA1"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

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


namespace WPF_Roots {
    
    
    /// <summary>
    /// PromptBox
    /// </summary>
    public partial class PromptBox : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 14 "..\..\PromptBox.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBlock tb_title;
        
        #line default
        #line hidden
        
        
        #line 19 "..\..\PromptBox.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Image i_icon;
        
        #line default
        #line hidden
        
        
        #line 23 "..\..\PromptBox.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBlock tb_msg;
        
        #line default
        #line hidden
        
        
        #line 25 "..\..\PromptBox.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ContentControl cc_msg;
        
        #line default
        #line hidden
        
        
        #line 27 "..\..\PromptBox.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DockPanel frame_buttons;
        
        #line default
        #line hidden
        
        
        #line 28 "..\..\PromptBox.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button bt_Ok;
        
        #line default
        #line hidden
        
        
        #line 29 "..\..\PromptBox.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button bt_Cancel;
        
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
            System.Uri resourceLocater = new System.Uri("/WPF_Roots;component/promptbox.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\PromptBox.xaml"
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
            
            #line 6 "..\..\PromptBox.xaml"
            ((WPF_Roots.PromptBox)(target)).Loaded += new System.Windows.RoutedEventHandler(this.Window_Loaded);
            
            #line default
            #line hidden
            
            #line 7 "..\..\PromptBox.xaml"
            ((WPF_Roots.PromptBox)(target)).PreviewKeyDown += new System.Windows.Input.KeyEventHandler(this.Window_PreviewKeyDown);
            
            #line default
            #line hidden
            return;
            case 2:
            this.tb_title = ((System.Windows.Controls.TextBlock)(target));
            return;
            case 3:
            this.i_icon = ((System.Windows.Controls.Image)(target));
            return;
            case 4:
            this.tb_msg = ((System.Windows.Controls.TextBlock)(target));
            return;
            case 5:
            this.cc_msg = ((System.Windows.Controls.ContentControl)(target));
            return;
            case 6:
            this.frame_buttons = ((System.Windows.Controls.DockPanel)(target));
            return;
            case 7:
            this.bt_Ok = ((System.Windows.Controls.Button)(target));
            
            #line 28 "..\..\PromptBox.xaml"
            this.bt_Ok.Click += new System.Windows.RoutedEventHandler(this.bt_Ok_Click);
            
            #line default
            #line hidden
            return;
            case 8:
            this.bt_Cancel = ((System.Windows.Controls.Button)(target));
            return;
            }
            this._contentLoaded = true;
        }
    }
}

