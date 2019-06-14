



using Microsoft.Office.Tools.Ribbon;


namespace R_TEx1316
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            RibbonToggleButton checkBox = Globals.Factory.GetRibbonFactory().CreateRibbonToggleButton();
            checkBox.Label = "User1";
            RibbonToggleButton checkBox2 = Globals.Factory.GetRibbonFactory().CreateRibbonToggleButton();
            checkBox2.Label = "User2";
            RibbonToggleButton checkBox3 = Globals.Factory.GetRibbonFactory().CreateRibbonToggleButton();
            checkBox3.Label = "User3";
            menu1.Label = "Users with access";
            menu1.Items.Add(checkBox);
            menu1.Items.Add(checkBox2);
            menu1.Items.Add(checkBox3);
            checkBox.Click += CheckBox_Click;
        }

        private void CheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.button3 = this.Factory.CreateRibbonButton();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.box2 = this.Factory.CreateRibbonBox();
            this.editBox7 = this.Factory.CreateRibbonEditBox();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.box4 = this.Factory.CreateRibbonBox();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.editBox5 = this.Factory.CreateRibbonEditBox();
            this.editBox4 = this.Factory.CreateRibbonEditBox();
            this.box5 = this.Factory.CreateRibbonBox();
            this.comboBox2 = this.Factory.CreateRibbonComboBox();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.editBox6 = this.Factory.CreateRibbonEditBox();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.box2.SuspendLayout();
            this.group4.SuspendLayout();
            this.box4.SuspendLayout();
            this.box5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "R-TEx";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.box1);
            this.group1.Label = "Login";
            this.group1.Name = "group1";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.editBox2);
            this.box1.Items.Add(this.editBox1);
            this.box1.Items.Add(this.checkBox1);
            this.box1.Items.Add(this.button1);
            this.box1.Name = "box1";
            // 
            // editBox2
            // 
            this.editBox2.Label = "E-mail";
            this.editBox2.Name = "editBox2";
            this.editBox2.Text = null;
            // 
            // editBox1
            // 
            this.editBox1.Label = "Password";
            this.editBox1.Name = "editBox1";
            this.editBox1.Text = null;
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "Remember Info";
            this.checkBox1.Name = "checkBox1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::R_TEx1316.Properties.Resources.checkmark_symbol_png_background_12;
            this.button1.Label = "Login";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.dropDown1);
            this.group3.Items.Add(this.button3);
            this.group3.Items.Add(this.menu2);
            this.group3.Label = "Cloud Workbooks";
            this.group3.Name = "group3";
            // 
            // dropDown1
            // 
            this.dropDown1.Label = "Workbooks:";
            this.dropDown1.Name = "dropDown1";
            // 
            // button3
            // 
            this.button3.Label = "Open Workbook";
            this.button3.Name = "button3";
            // 
            // menu2
            // 
            this.menu2.Label = "Users With Access";
            this.menu2.Name = "menu2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.box2);
            this.group2.Label = "New Cloud Workbook";
            this.group2.Name = "group2";
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.menu1);
            this.box2.Items.Add(this.editBox7);
            this.box2.Items.Add(this.button2);
            this.box2.Name = "box2";
            // 
            // editBox7
            // 
            this.editBox7.Label = "Workbook Name";
            this.editBox7.Name = "editBox7";
            this.editBox7.Text = null;
            // 
            // button2
            // 
            this.button2.Label = "Create Cloud Workbook";
            this.button2.Name = "button2";
            // 
            // group4
            // 
            this.group4.Items.Add(this.box4);
            this.group4.Items.Add(this.box5);
            this.group4.Label = "User Management";
            this.group4.Name = "group4";
            // 
            // box4
            // 
            this.box4.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box4.Items.Add(this.comboBox1);
            this.box4.Items.Add(this.editBox5);
            this.box4.Items.Add(this.editBox4);
            this.box4.Name = "box4";
            // 
            // comboBox1
            // 
            ribbonDropDownItemImpl1.Label = "Add User";
            ribbonDropDownItemImpl2.Label = "Edit User";
            this.comboBox1.Items.Add(ribbonDropDownItemImpl1);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl2);
            this.comboBox1.Label = "Action";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Text = null;
            // 
            // editBox5
            // 
            this.editBox5.Label = "First Name";
            this.editBox5.Name = "editBox5";
            this.editBox5.Text = null;
            this.editBox5.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox5_TextChanged);
            // 
            // editBox4
            // 
            this.editBox4.Label = "Last Name";
            this.editBox4.Name = "editBox4";
            this.editBox4.Text = null;
            // 
            // box5
            // 
            this.box5.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box5.Items.Add(this.comboBox2);
            this.box5.Items.Add(this.editBox3);
            this.box5.Items.Add(this.editBox6);
            this.box5.Name = "box5";
            // 
            // comboBox2
            // 
            this.comboBox2.Label = "User To Edit";
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Text = null;
            // 
            // editBox3
            // 
            this.editBox3.Label = "  E-mail";
            this.editBox3.Name = "editBox3";
            this.editBox3.Text = null;
            // 
            // editBox6
            // 
            this.editBox6.Label = "  Password";
            this.editBox6.Name = "editBox6";
            this.editBox6.Text = null;
            // 
            // menu1
            // 
            this.menu1.Label = "Users With Access";
            this.menu1.Name = "menu1";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal RibbonGroup group3;
        internal RibbonBox box2;
        internal RibbonButton button2;
        internal RibbonGroup group4;
        internal RibbonComboBox comboBox1;
        internal RibbonBox box4;
        internal RibbonEditBox editBox5;
        internal RibbonEditBox editBox4;
        internal RibbonBox box5;
        internal RibbonEditBox editBox3;
        internal RibbonEditBox editBox6;
        internal RibbonComboBox comboBox2;
        internal RibbonDropDown dropDown1;
        internal RibbonButton button3;
        internal RibbonCheckBox checkBox1;
        internal RibbonMenu menu2;
        internal RibbonEditBox editBox7;
        internal RibbonMenu menu1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
