﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

#pragma warning disable 414
namespace WordDocumentGenerator.WordRefreshableDocumentAddin {
    
    
    /// 
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class ThisDocument : Microsoft.Office.Tools.Word.DocumentBase {
        
        internal Microsoft.Office.Tools.ActionsPane ActionsPane;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl plainTextContentControl1;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl plainTextContentControl2;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl plainTextContentControl3;
        
        internal Microsoft.Office.Tools.Word.RichTextContentControl richTextContentControl1;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl plainTextContentControl4;
        
        internal Microsoft.Office.Tools.Word.RichTextContentControl richTextContentControl2;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl plainTextContentControl5;
        
        internal Microsoft.Office.Tools.Word.RichTextContentControl richTextContentControl3;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl plainTextContentControl6;
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        private global::System.Object missing = global::System.Type.Missing;
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        internal Microsoft.Office.Interop.Word.Application ThisApplication;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public ThisDocument(global::Microsoft.Office.Tools.Word.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "ThisDocument", "ThisDocument") {
            Globals.Factory = factory;
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void Initialize() {
            base.Initialize();
            this.ThisApplication = this.GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), "Application");
            Globals.ThisDocument = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            this.InternalStartup();
            this.OnStartup();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void InitializeDataBindings() {
            this.BeginInitialization();
            this.BindToData();
            this.EndInitialization();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeCachedData() {
            if ((this.DataHost == null)) {
                return;
            }
            if (this.DataHost.IsCacheInitialized) {
                this.DataHost.FillCachedData(this);
            }
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BindToData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StartCaching(string MemberName) {
            this.DataHost.StartCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StopCaching(string MemberName) {
            this.DataHost.StopCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool IsCached(string MemberName) {
            return this.DataHost.IsCached(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BeginInitialization() {
            this.BeginInit();
            this.ActionsPane.BeginInit();
            this.plainTextContentControl1.BeginInit();
            this.plainTextContentControl2.BeginInit();
            this.plainTextContentControl3.BeginInit();
            this.richTextContentControl1.BeginInit();
            this.plainTextContentControl4.BeginInit();
            this.richTextContentControl2.BeginInit();
            this.plainTextContentControl5.BeginInit();
            this.richTextContentControl3.BeginInit();
            this.plainTextContentControl6.BeginInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void EndInitialization() {
            this.plainTextContentControl6.EndInit();
            this.richTextContentControl3.EndInit();
            this.plainTextContentControl5.EndInit();
            this.richTextContentControl2.EndInit();
            this.plainTextContentControl4.EndInit();
            this.richTextContentControl1.EndInit();
            this.plainTextContentControl3.EndInit();
            this.plainTextContentControl2.EndInit();
            this.plainTextContentControl1.EndInit();
            this.ActionsPane.EndInit();
            this.EndInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeControls() {
            this.ActionsPane = Globals.Factory.CreateActionsPane(null, null, "ActionsPane", "ActionsPane", this);
            this.plainTextContentControl1 = Globals.Factory.CreatePlainTextContentControl(null, null, "397949701", "plainTextContentControl1", this);
            this.plainTextContentControl2 = Globals.Factory.CreatePlainTextContentControl(null, null, "3383654310", "plainTextContentControl2", this);
            this.plainTextContentControl3 = Globals.Factory.CreatePlainTextContentControl(null, null, "954136284", "plainTextContentControl3", this);
            this.richTextContentControl1 = Globals.Factory.CreateRichTextContentControl(null, null, "467246042", "richTextContentControl1", this);
            this.plainTextContentControl4 = Globals.Factory.CreatePlainTextContentControl(null, null, "2050569460", "plainTextContentControl4", this);
            this.richTextContentControl2 = Globals.Factory.CreateRichTextContentControl(null, null, "2299482968", "richTextContentControl2", this);
            this.plainTextContentControl5 = Globals.Factory.CreatePlainTextContentControl(null, null, "2946110922", "plainTextContentControl5", this);
            this.richTextContentControl3 = Globals.Factory.CreateRichTextContentControl(null, null, "3758225702", "richTextContentControl3", this);
            this.plainTextContentControl6 = Globals.Factory.CreatePlainTextContentControl(null, null, "2672279940", "plainTextContentControl6", this);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeComponents() {
            // 
            // ActionsPane
            // 
            this.ActionsPane.AutoSize = false;
            this.ActionsPane.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            // 
            // plainTextContentControl1
            // 
            this.plainTextContentControl1.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // plainTextContentControl2
            // 
            this.plainTextContentControl2.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // plainTextContentControl3
            // 
            this.plainTextContentControl3.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // richTextContentControl1
            // 
            this.richTextContentControl1.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // plainTextContentControl4
            // 
            this.plainTextContentControl4.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // richTextContentControl2
            // 
            this.richTextContentControl2.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // plainTextContentControl5
            // 
            this.plainTextContentControl5.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // richTextContentControl3
            // 
            this.richTextContentControl3.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // plainTextContentControl6
            // 
            this.plainTextContentControl6.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // ThisDocument
            // 
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName) {
            return this.DataHost.NeedsFill(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void OnShutdown() {
            this.plainTextContentControl6.Dispose();
            this.richTextContentControl3.Dispose();
            this.plainTextContentControl5.Dispose();
            this.richTextContentControl2.Dispose();
            this.plainTextContentControl4.Dispose();
            this.richTextContentControl1.Dispose();
            this.plainTextContentControl3.Dispose();
            this.plainTextContentControl2.Dispose();
            this.plainTextContentControl1.Dispose();
            this.ActionsPane.Dispose();
            base.OnShutdown();
        }
    }
    
    /// 
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
    internal sealed partial class Globals {
        
        /// 
        private Globals() {
        }
        
        private static ThisDocument _ThisDocument;
        
        private static global::Microsoft.Office.Tools.Word.Factory _factory;
        
        private static ThisRibbonCollection _ThisRibbonCollection;
        
        internal static ThisDocument ThisDocument {
            get {
                return _ThisDocument;
            }
            set {
                if ((_ThisDocument == null)) {
                    _ThisDocument = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
        
        internal static global::Microsoft.Office.Tools.Word.Factory Factory {
            get {
                return _factory;
            }
            set {
                if ((_factory == null)) {
                    _factory = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
        
        internal static ThisRibbonCollection Ribbons {
            get {
                if ((_ThisRibbonCollection == null)) {
                    _ThisRibbonCollection = new ThisRibbonCollection(_factory.GetRibbonFactory());
                }
                return _ThisRibbonCollection;
            }
        }
    }
    
    /// 
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
    internal sealed partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonCollectionBase {
        
        /// 
        internal ThisRibbonCollection(global::Microsoft.Office.Tools.Ribbon.RibbonFactory factory) : 
                base(factory) {
        }
    }
}