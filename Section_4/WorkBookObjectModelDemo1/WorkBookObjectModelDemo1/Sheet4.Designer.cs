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
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(4)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class Sheet4 : Microsoft.Office.Tools.Excel.WorksheetBase {
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        private global::System.Object missing = global::System.Type.Missing;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnR1C1Notation;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnActiveCellApp;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnSheetObjectCell;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnReadValue;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnSetValue;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnSetValueCont;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnLoopCells;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnCellFormat;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnFormatWithStyle;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public Sheet4(global::Microsoft.Office.Tools.Excel.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "Sheet4", "Sheet4") {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void Initialize() {
            base.Initialize();
            Globals.Sheet4 = this;
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
            this.btnR1C1Notation = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "22E21BBC924BF62462A29DE924EE1090D9E092", "22E21BBC924BF62462A29DE924EE1090D9E092", this, "btnR1C1Notation");
            this.btnActiveCellApp = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "36181A4433C600346F03BDC337F9495E68A2E3", "36181A4433C600346F03BDC337F9495E68A2E3", this, "btnActiveCellApp");
            this.btnSheetObjectCell = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "42FEF61344BCD844C4C497224D4BC67014FAB4", "42FEF61344BCD844C4C497224D4BC67014FAB4", this, "btnSheetObjectCell");
            this.btnReadValue = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "5A6584FB25DE5354DA958A285CE8E01FC8FAF5", "5A6584FB25DE5354DA958A285CE8E01FC8FAF5", this, "btnReadValue");
            this.btnSetValue = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "6297D2C656AB8F645E169B1466467835B828B6", "6297D2C656AB8F645E169B1466467835B828B6", this, "btnSetValue");
            this.btnSetValueCont = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "7C3148B287F4CF7475C79D4077719A5E4F6277", "7C3148B287F4CF7475C79D4077719A5E4F6277", this, "btnSetValueCont");
            this.btnLoopCells = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "8631245318678984CA08BD888B7176933F67E8", "8631245318678984CA08BD888B7176933F67E8", this, "btnLoopCells");
            this.btnCellFormat = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "96F71D6E0922FD940AB9A6269EF439A3FBC279", "96F71D6E0922FD940AB9A6269EF439A3FBC279", this, "btnCellFormat");
            this.btnFormatWithStyle = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "1A9FA3C5313875148381BFC01B8615A5007A21", "1A9FA3C5313875148381BFC01B8615A5007A21", this, "btnFormatWithStyle");
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeComponents() {
            // 
            // btnR1C1Notation
            // 
            this.btnR1C1Notation.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnR1C1Notation.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnR1C1Notation.Name = "btnR1C1Notation";
            this.btnR1C1Notation.Text = "R1C1 Notation";
            this.btnR1C1Notation.UseVisualStyleBackColor = false;
            // 
            // btnActiveCellApp
            // 
            this.btnActiveCellApp.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnActiveCellApp.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnActiveCellApp.Name = "btnActiveCellApp";
            this.btnActiveCellApp.Text = "Active Cell";
            this.btnActiveCellApp.UseVisualStyleBackColor = false;
            // 
            // btnSheetObjectCell
            // 
            this.btnSheetObjectCell.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnSheetObjectCell.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnSheetObjectCell.Name = "btnSheetObjectCell";
            this.btnSheetObjectCell.Text = "Sheet Object Cell";
            this.btnSheetObjectCell.UseVisualStyleBackColor = false;
            // 
            // btnReadValue
            // 
            this.btnReadValue.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnReadValue.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnReadValue.Name = "btnReadValue";
            this.btnReadValue.Text = "Read Value";
            this.btnReadValue.UseVisualStyleBackColor = false;
            // 
            // btnSetValue
            // 
            this.btnSetValue.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnSetValue.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnSetValue.Name = "btnSetValue";
            this.btnSetValue.Text = "Set Cell Value";
            this.btnSetValue.UseVisualStyleBackColor = false;
            // 
            // btnSetValueCont
            // 
            this.btnSetValueCont.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnSetValueCont.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnSetValueCont.Name = "btnSetValueCont";
            this.btnSetValueCont.Text = "Set Values Contiguous";
            this.btnSetValueCont.UseVisualStyleBackColor = false;
            // 
            // btnLoopCells
            // 
            this.btnLoopCells.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnLoopCells.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnLoopCells.Name = "btnLoopCells";
            this.btnLoopCells.Text = "Loop Cells";
            this.btnLoopCells.UseVisualStyleBackColor = false;
            // 
            // btnCellFormat
            // 
            this.btnCellFormat.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnCellFormat.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnCellFormat.Name = "btnCellFormat";
            this.btnCellFormat.Text = " Cells Format";
            this.btnCellFormat.UseVisualStyleBackColor = false;
            // 
            // btnFormatWithStyle
            // 
            this.btnFormatWithStyle.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnFormatWithStyle.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnFormatWithStyle.Name = "btnFormatWithStyle";
            this.btnFormatWithStyle.Text = "Using Style";
            this.btnFormatWithStyle.UseVisualStyleBackColor = false;
            // 
            // Sheet4
            // 
            this.btnR1C1Notation.BindingContext = this.BindingContext;
            this.btnActiveCellApp.BindingContext = this.BindingContext;
            this.btnSheetObjectCell.BindingContext = this.BindingContext;
            this.btnReadValue.BindingContext = this.BindingContext;
            this.btnSetValue.BindingContext = this.BindingContext;
            this.btnSetValueCont.BindingContext = this.BindingContext;
            this.btnLoopCells.BindingContext = this.BindingContext;
            this.btnCellFormat.BindingContext = this.BindingContext;
            this.btnFormatWithStyle.BindingContext = this.BindingContext;
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName) {
            return this.DataHost.NeedsFill(this, MemberName);
        }
    }
    
    internal sealed partial class Globals {
        
        private static Sheet4 _Sheet4;
        
        internal static Sheet4 Sheet4 {
            get {
                return _Sheet4;
            }
            set {
                if ((_Sheet4 == null)) {
                    _Sheet4 = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
    }
}