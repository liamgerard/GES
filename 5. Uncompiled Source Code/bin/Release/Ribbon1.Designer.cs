namespace GES
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
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
            this.GES = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.openSC = this.Factory.CreateRibbonButton();
            this.openForm = this.Factory.CreateRibbonButton();
            this.openDK = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.restartAddIn = this.Factory.CreateRibbonButton();
            this.enableAddIn = this.Factory.CreateRibbonButton();
            this.disableAddIn = this.Factory.CreateRibbonButton();
            this.Functions = this.Factory.CreateRibbonGroup();
            this.buttonGroup5 = this.Factory.CreateRibbonButtonGroup();
            this.genFormButton = this.Factory.CreateRibbonButton();
            this.perButton = this.Factory.CreateRibbonButton();
            this.multButton = this.Factory.CreateRibbonButton();
            this.buttonGroup4 = this.Factory.CreateRibbonButtonGroup();
            this.currButton = this.Factory.CreateRibbonButton();
            this.fCurrButton = this.Factory.CreateRibbonButton();
            this.BinButton = this.Factory.CreateRibbonButton();
            this.buttonGroup3 = this.Factory.CreateRibbonButtonGroup();
            this.incDecButton = this.Factory.CreateRibbonButton();
            this.decDecButton = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.toggNegButton = this.Factory.CreateRibbonButton();
            this.shiftDecLButton = this.Factory.CreateRibbonButton();
            this.shiftDecRButton = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.version = this.Factory.CreateRibbonLabel();
            this.GES.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.Functions.SuspendLayout();
            this.buttonGroup5.SuspendLayout();
            this.buttonGroup4.SuspendLayout();
            this.buttonGroup3.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // GES
            // 
            this.GES.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.GES.Groups.Add(this.group2);
            this.GES.Groups.Add(this.group1);
            this.GES.Groups.Add(this.Functions);
            this.GES.Groups.Add(this.group3);
            this.GES.Groups.Add(this.group4);
            this.GES.Label = "GES";
            this.GES.Name = "GES";
            // 
            // group2
            // 
            this.group2.Items.Add(this.openSC);
            this.group2.Items.Add(this.openForm);
            this.group2.Items.Add(this.openDK);
            this.group2.Label = "Add-In Settings";
            this.group2.Name = "group2";
            // 
            // openSC
            // 
            this.openSC.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.openSC.KeyTip = "S";
            this.openSC.Label = "Shortcuts Menu";
            this.openSC.Name = "openSC";
            this.openSC.OfficeImageId = "PropertySheet";
            this.openSC.ShowImage = true;
            this.openSC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openSC_Click);
            // 
            // openForm
            // 
            this.openForm.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.openForm.KeyTip = "T";
            this.openForm.Label = "Formatting Menu";
            this.openForm.Name = "openForm";
            this.openForm.OfficeImageId = "PropertySheet";
            this.openForm.ShowImage = true;
            this.openForm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openForm_Click);
            // 
            // openDK
            // 
            this.openDK.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.openDK.KeyTip = "K";
            this.openDK.Label = "Disabled Keys Menu";
            this.openDK.Name = "openDK";
            this.openDK.OfficeImageId = "PropertySheet";
            this.openDK.ShowImage = true;
            this.openDK.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openDK_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.restartAddIn);
            this.group1.Items.Add(this.enableAddIn);
            this.group1.Items.Add(this.disableAddIn);
            this.group1.Label = "Manage Add-In";
            this.group1.Name = "group1";
            // 
            // restartAddIn
            // 
            this.restartAddIn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.restartAddIn.KeyTip = "R";
            this.restartAddIn.Label = "Restart Add-In";
            this.restartAddIn.Name = "restartAddIn";
            this.restartAddIn.OfficeImageId = "Refresh";
            this.restartAddIn.ShowImage = true;
            this.restartAddIn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.restartAddIn_Click);
            // 
            // enableAddIn
            // 
            this.enableAddIn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.enableAddIn.KeyTip = "E";
            this.enableAddIn.Label = "Enable Add-In";
            this.enableAddIn.Name = "enableAddIn";
            this.enableAddIn.OfficeImageId = "MacroPlay";
            this.enableAddIn.ShowImage = true;
            this.enableAddIn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.enableAddIn_Click);
            // 
            // disableAddIn
            // 
            this.disableAddIn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.disableAddIn.KeyTip = "D";
            this.disableAddIn.Label = "Disable Add-In";
            this.disableAddIn.Name = "disableAddIn";
            this.disableAddIn.OfficeImageId = "MacroRecorderStop";
            this.disableAddIn.ShowImage = true;
            this.disableAddIn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.disableAddIn_Click);
            // 
            // Functions
            // 
            this.Functions.Items.Add(this.buttonGroup5);
            this.Functions.Items.Add(this.buttonGroup4);
            this.Functions.Items.Add(this.buttonGroup3);
            this.Functions.Label = "Number Formatting";
            this.Functions.Name = "Functions";
            // 
            // buttonGroup5
            // 
            this.buttonGroup5.Items.Add(this.genFormButton);
            this.buttonGroup5.Items.Add(this.perButton);
            this.buttonGroup5.Items.Add(this.multButton);
            this.buttonGroup5.Name = "buttonGroup5";
            // 
            // genFormButton
            // 
            this.genFormButton.Label = "General Formats";
            this.genFormButton.Name = "genFormButton";
            this.genFormButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.genFormButton_Click);
            // 
            // perButton
            // 
            this.perButton.Label = "Percentages";
            this.perButton.Name = "perButton";
            this.perButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.perButton_Click);
            // 
            // multButton
            // 
            this.multButton.Label = "Multiples";
            this.multButton.Name = "multButton";
            this.multButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.multButton_Click);
            // 
            // buttonGroup4
            // 
            this.buttonGroup4.Items.Add(this.currButton);
            this.buttonGroup4.Items.Add(this.fCurrButton);
            this.buttonGroup4.Items.Add(this.BinButton);
            this.buttonGroup4.Name = "buttonGroup4";
            // 
            // currButton
            // 
            this.currButton.Label = "Currencies";
            this.currButton.Name = "currButton";
            this.currButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.currButton_Click);
            // 
            // fCurrButton
            // 
            this.fCurrButton.Label = "Foreign Currencies";
            this.fCurrButton.Name = "fCurrButton";
            this.fCurrButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fCurrButton_Click);
            // 
            // BinButton
            // 
            this.BinButton.Label = "Binaries";
            this.BinButton.Name = "BinButton";
            this.BinButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BinButton_Click);
            // 
            // buttonGroup3
            // 
            this.buttonGroup3.Items.Add(this.incDecButton);
            this.buttonGroup3.Items.Add(this.decDecButton);
            this.buttonGroup3.Name = "buttonGroup3";
            // 
            // incDecButton
            // 
            this.incDecButton.Label = "Increase Decimal";
            this.incDecButton.Name = "incDecButton";
            this.incDecButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.incDecButton_Click);
            // 
            // decDecButton
            // 
            this.decDecButton.Label = "Decrease Decimal";
            this.decDecButton.Name = "decDecButton";
            this.decDecButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.decDecButton_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.toggNegButton);
            this.group3.Items.Add(this.shiftDecLButton);
            this.group3.Items.Add(this.shiftDecRButton);
            this.group3.Label = "Formula Editing";
            this.group3.Name = "group3";
            // 
            // toggNegButton
            // 
            this.toggNegButton.Label = "Toggle Negative";
            this.toggNegButton.Name = "toggNegButton";
            this.toggNegButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggNegButton_Click);
            // 
            // shiftDecLButton
            // 
            this.shiftDecLButton.Label = "Shift Decimal Left";
            this.shiftDecLButton.Name = "shiftDecLButton";
            this.shiftDecLButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.shiftDecLButton_Click);
            // 
            // shiftDecRButton
            // 
            this.shiftDecRButton.Label = "Shift Decimal Right";
            this.shiftDecRButton.Name = "shiftDecRButton";
            this.shiftDecRButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.shiftDecRButton_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.label1);
            this.group4.Items.Add(this.label2);
            this.group4.Items.Add(this.version);
            this.group4.Label = "Developer";
            this.group4.Name = "group4";
            // 
            // label1
            // 
            this.label1.Label = "Developed by Liam Gerard";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "Gerard Excel Shortcuts (GES)";
            this.label2.Name = "label2";
            // 
            // version
            // 
            this.version.Label = "Demo Version";
            this.version.Name = "version";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.GES);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.GES.ResumeLayout(false);
            this.GES.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Functions.ResumeLayout(false);
            this.Functions.PerformLayout();
            this.buttonGroup5.ResumeLayout(false);
            this.buttonGroup5.PerformLayout();
            this.buttonGroup4.ResumeLayout(false);
            this.buttonGroup4.PerformLayout();
            this.buttonGroup3.ResumeLayout(false);
            this.buttonGroup3.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab GES;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Functions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openSC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openForm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton restartAddIn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton genFormButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton currButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fCurrButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton perButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton multButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BinButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toggNegButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton shiftDecLButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton shiftDecRButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton incDecButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel version;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton decDecButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openDK;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton enableAddIn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton disableAddIn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
