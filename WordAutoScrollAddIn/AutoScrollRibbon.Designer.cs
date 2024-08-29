namespace WordAutoScrollAddIn
{
    partial class AutoScrollRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public AutoScrollRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.AutoScrollGroup = this.Factory.CreateRibbonGroup();
            this.StartScrollButton = this.Factory.CreateRibbonButton();
            this.StopScrollButton = this.Factory.CreateRibbonButton();
            this.LinesEditBox = this.Factory.CreateRibbonEditBox();
            this.DelayEditBox = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.AutoScrollGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.AutoScrollGroup);
            this.tab1.Label = "自动滚动插件";
            this.tab1.Name = "tab1";
            // 
            // AutoScrollGroup
            // 
            this.AutoScrollGroup.Items.Add(this.StartScrollButton);
            this.AutoScrollGroup.Items.Add(this.StopScrollButton);
            this.AutoScrollGroup.Items.Add(this.LinesEditBox);
            this.AutoScrollGroup.Items.Add(this.DelayEditBox);
            this.AutoScrollGroup.Label = "自动滚动";
            this.AutoScrollGroup.Name = "AutoScrollGroup";
            // 
            // StartScrollButton
            // 
            this.StartScrollButton.Label = "开始滚动";
            this.StartScrollButton.Name = "StartScrollButton";
            this.StartScrollButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StartScrollButton_Click);
            // 
            // StopScrollButton
            // 
            this.StopScrollButton.Label = "停止滚动";
            this.StopScrollButton.Name = "StopScrollButton";
            this.StopScrollButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StopScrollButton_Click);
            // 
            // LinesEditBox
            // 
            this.LinesEditBox.Label = "滚动行数";
            this.LinesEditBox.Name = "LinesEditBox";
            this.LinesEditBox.Text = "1";
            this.LinesEditBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LinesEditBox_TextChanged);
            // 
            // DelayEditBox
            // 
            this.DelayEditBox.Label = "停留时间(ms)";
            this.DelayEditBox.Name = "DelayEditBox";
            this.DelayEditBox.Text = "100";
            this.DelayEditBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DelayEditBox_TextChanged);
            // 
            // AutoScrollRibbon
            // 
            this.Name = "AutoScrollRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AutoScrollRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.AutoScrollGroup.ResumeLayout(false);
            this.AutoScrollGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AutoScrollGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton StartScrollButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton StopScrollButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox LinesEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox DelayEditBox;
    }

    partial class ThisRibbonCollection
    {
        internal AutoScrollRibbon AutoScrollRibbon
        {
            get { return this.GetRibbon<AutoScrollRibbon>(); }
        }
    }
}
