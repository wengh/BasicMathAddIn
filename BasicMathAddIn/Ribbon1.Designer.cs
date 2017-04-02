namespace BasicMathAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1 ()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose (bool disposing)
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
        private void InitializeComponent ()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.checkBox2 = this.Factory.CreateRibbonCheckBox();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.editBox4 = this.Factory.CreateRibbonEditBox();
            this.editBox5 = this.Factory.CreateRibbonEditBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.editBox6 = this.Factory.CreateRibbonEditBox();
            this.editBox7 = this.Factory.CreateRibbonEditBox();
            this.checkBox3 = this.Factory.CreateRibbonCheckBox();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.editBox8 = this.Factory.CreateRibbonEditBox();
            this.editBox9 = this.Factory.CreateRibbonEditBox();
            this.checkBox4 = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            resources.ApplyResources(this.tab1, "tab1");
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Items.Add(this.editBox2);
            this.group2.Items.Add(this.checkBox1);
            this.group2.Items.Add(this.checkBox2);
            this.group2.Items.Add(this.dropDown1);
            this.group2.Items.Add(this.checkBox3);
            this.group2.Items.Add(this.checkBox4);
            resources.ApplyResources(this.group2, "group2");
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::BasicMathAddIn.Properties.Resources.Dice_80px;
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // editBox2
            // 
            resources.ApplyResources(this.editBox2, "editBox2");
            this.editBox2.Name = "editBox2";
            // 
            // checkBox1
            // 
            this.checkBox1.Checked = true;
            resources.ApplyResources(this.checkBox1, "checkBox1");
            this.checkBox1.Name = "checkBox1";
            // 
            // checkBox2
            // 
            resources.ApplyResources(this.checkBox2, "checkBox2");
            this.checkBox2.Name = "checkBox2";
            // 
            // dropDown1
            // 
            resources.ApplyResources(ribbonDropDownItemImpl1, "ribbonDropDownItemImpl1");
            resources.ApplyResources(ribbonDropDownItemImpl2, "ribbonDropDownItemImpl2");
            this.dropDown1.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl2);
            resources.ApplyResources(this.dropDown1, "dropDown1");
            this.dropDown1.Name = "dropDown1";
            this.dropDown1.ShowItemImage = false;
            // 
            // group1
            // 
            this.group1.Items.Add(this.editBox1);
            this.group1.Items.Add(this.editBox3);
            resources.ApplyResources(this.group1, "group1");
            this.group1.Name = "group1";
            // 
            // editBox1
            // 
            resources.ApplyResources(this.editBox1, "editBox1");
            this.editBox1.Name = "editBox1";
            // 
            // editBox3
            // 
            resources.ApplyResources(this.editBox3, "editBox3");
            this.editBox3.Name = "editBox3";
            // 
            // group3
            // 
            this.group3.Items.Add(this.editBox4);
            this.group3.Items.Add(this.editBox5);
            resources.ApplyResources(this.group3, "group3");
            this.group3.Name = "group3";
            // 
            // editBox4
            // 
            resources.ApplyResources(this.editBox4, "editBox4");
            this.editBox4.Name = "editBox4";
            // 
            // editBox5
            // 
            resources.ApplyResources(this.editBox5, "editBox5");
            this.editBox5.Name = "editBox5";
            // 
            // group4
            // 
            this.group4.Items.Add(this.editBox6);
            this.group4.Items.Add(this.editBox7);
            resources.ApplyResources(this.group4, "group4");
            this.group4.Name = "group4";
            // 
            // editBox6
            // 
            resources.ApplyResources(this.editBox6, "editBox6");
            this.editBox6.Name = "editBox6";
            // 
            // editBox7
            // 
            resources.ApplyResources(this.editBox7, "editBox7");
            this.editBox7.Name = "editBox7";
            // 
            // checkBox3
            // 
            resources.ApplyResources(this.checkBox3, "checkBox3");
            this.checkBox3.Name = "checkBox3";
            // 
            // group5
            // 
            this.group5.Items.Add(this.editBox8);
            this.group5.Items.Add(this.editBox9);
            resources.ApplyResources(this.group5, "group5");
            this.group5.Name = "group5";
            // 
            // editBox8
            // 
            resources.ApplyResources(this.editBox8, "editBox8");
            this.editBox8.Name = "editBox8";
            // 
            // editBox9
            // 
            resources.ApplyResources(this.editBox9, "editBox9");
            this.editBox9.Name = "editBox9";
            // 
            // checkBox4
            // 
            resources.ApplyResources(this.checkBox4, "checkBox4");
            this.checkBox4.Name = "checkBox4";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox5;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox6;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox7;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox8;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox9;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox4;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
