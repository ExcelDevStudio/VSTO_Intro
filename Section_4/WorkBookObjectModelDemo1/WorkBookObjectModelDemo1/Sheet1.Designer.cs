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
namespace WorkBookObjectModelDemo1 {
    
    
    /// 
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(1)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class Sheet1 : Microsoft.Office.Tools.Excel.WorksheetBase {
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        private global::System.Object missing = global::System.Type.Missing;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnSetStatusMsg;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnClearStatusMsg;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnWorkSheetFunction;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnCalculateSheet;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnReCalculateAll;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public Sheet1(global::Microsoft.Office.Tools.Excel.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "Sheet1", "Sheet1") {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void Initialize() {
            base.Initialize();
            Globals.Sheet1 = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            this.InternalStartup();
            this.OnStartup();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void InitializeDataBindings() {
            this.BeginInitialization();
            this.BindToData();
            this.EndInitialization();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
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
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
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
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BeginInitialization() {
            this.BeginInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void EndInitialization() {
            this.EndInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeControls() {
            this.btnSetStatusMsg = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "1C775D7A51E3F41486018FF61C51C5B78C6EF1", "1C775D7A51E3F41486018FF61C51C5B78C6EF1", this, "btnSetStatusMsg");
            this.btnClearStatusMsg = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "338AFC9EC3FFA334D983B76C3DA6DD798161B3", "338AFC9EC3FFA334D983B76C3DA6DD798161B3", this, "btnClearStatusMsg");
            this.btnWorkSheetFunction = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "65363583461EBC64114695B1622E4C7234B6A6", "65363583461EBC64114695B1622E4C7234B6A6", this, "btnWorkSheetFunction");
            this.btnCalculateSheet = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "7C20576B17592C746A679CD77C627266454237", "7C20576B17592C746A679CD77C627266454237", this, "btnCalculateSheet");
            this.btnReCalculateAll = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "93C5D7106919789435D9862195281E9AFDBCF9", "93C5D7106919789435D9862195281E9AFDBCF9", this, "btnReCalculateAll");
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeComponents() {
            // 
            // btnSetStatusMsg
            // 
            this.btnSetStatusMsg.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnSetStatusMsg.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnSetStatusMsg.Name = "btnSetStatusMsg";
            this.btnSetStatusMsg.Text = "Status Message";
            this.btnSetStatusMsg.UseVisualStyleBackColor = false;
            // 
            // btnClearStatusMsg
            // 
            this.btnClearStatusMsg.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnClearStatusMsg.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnClearStatusMsg.Name = "btnClearStatusMsg";
            this.btnClearStatusMsg.Text = "Clear Message";
            this.btnClearStatusMsg.UseVisualStyleBackColor = false;
            // 
            // btnWorkSheetFunction
            // 
            this.btnWorkSheetFunction.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnWorkSheetFunction.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnWorkSheetFunction.Name = "btnWorkSheetFunction";
            this.btnWorkSheetFunction.Text = "Worksheet Function";
            this.btnWorkSheetFunction.UseVisualStyleBackColor = false;
            // 
            // btnCalculateSheet
            // 
            this.btnCalculateSheet.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnCalculateSheet.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnCalculateSheet.Name = "btnCalculateSheet";
            this.btnCalculateSheet.Text = "Calculate Sheet";
            this.btnCalculateSheet.UseVisualStyleBackColor = false;
            // 
            // btnReCalculateAll
            // 
            this.btnReCalculateAll.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnReCalculateAll.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnReCalculateAll.Name = "btnReCalculateAll";
            this.btnReCalculateAll.Text = "Calculate All";
            this.btnReCalculateAll.UseVisualStyleBackColor = false;
            // 
            // Sheet1
            // 
            this.btnSetStatusMsg.BindingContext = this.BindingContext;
            this.btnClearStatusMsg.BindingContext = this.BindingContext;
            this.btnWorkSheetFunction.BindingContext = this.BindingContext;
            this.btnCalculateSheet.BindingContext = this.BindingContext;
            this.btnReCalculateAll.BindingContext = this.BindingContext;
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName) {
            return this.DataHost.NeedsFill(this, MemberName);
        }
    }
    
    internal sealed partial class Globals {
        
        private static Sheet1 _Sheet1;
        
        internal static Sheet1 Sheet1 {
            get {
                return _Sheet1;
            }
            set {
                if ((_Sheet1 == null)) {
                    _Sheet1 = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
    }
}