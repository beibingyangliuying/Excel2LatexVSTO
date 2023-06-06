namespace Excel2Latex
{
    partial class RibbonExcel2Latex : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonExcel2Latex()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tabExcel2Latex = this.Factory.CreateRibbonTab();
            this.groupTransfer = this.Factory.CreateRibbonGroup();
            this.buttonTransfer = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tabExcel2Latex.SuspendLayout();
            this.groupTransfer.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tabExcel2Latex
            // 
            this.tabExcel2Latex.Groups.Add(this.groupTransfer);
            this.tabExcel2Latex.Label = "Excel2Latex";
            this.tabExcel2Latex.Name = "tabExcel2Latex";
            // 
            // groupTransfer
            // 
            this.groupTransfer.Items.Add(this.buttonTransfer);
            this.groupTransfer.Label = "转换";
            this.groupTransfer.Name = "groupTransfer";
            // 
            // buttonTransfer
            // 
            this.buttonTransfer.Label = "选定区域并生成tex表格代码";
            this.buttonTransfer.Name = "buttonTransfer";
            this.buttonTransfer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonTransfer_Click);
            // 
            // RibbonExcel2Latex
            // 
            this.Name = "RibbonExcel2Latex";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabExcel2Latex);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonExcel2Latex_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabExcel2Latex.ResumeLayout(false);
            this.tabExcel2Latex.PerformLayout();
            this.groupTransfer.ResumeLayout(false);
            this.groupTransfer.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabExcel2Latex;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTransfer;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTransfer;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonExcel2Latex RibbonExcel2Latex
        {
            get { return this.GetRibbon<RibbonExcel2Latex>(); }
        }
    }
}
