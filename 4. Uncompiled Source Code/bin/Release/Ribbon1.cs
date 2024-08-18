using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;

namespace GES
{
    public partial class Ribbon1
    {
        public static Ribbon1 Instance { get; private set; }
        public void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Instance = this;
        }
        public void DisableAllControls()
        {
            this.openSC.Enabled = false;
            this.openForm.Enabled = false;
            this.openDK.Enabled = false;
            this.restartAddIn.Enabled = false;
            this.enableAddIn.Enabled = false;
            this.disableAddIn.Enabled = false;
            this.genFormButton.Enabled = false;
            this.currButton.Enabled = false;
            this.fCurrButton.Enabled = false;
            this.perButton.Enabled = false;
            this.multButton.Enabled = false;
            this.BinButton.Enabled = false;
            this.incDecButton.Enabled = false;
            this.decDecButton.Enabled = false;
            this.toggNegButton.Enabled = false;
            this.shiftDecLButton.Enabled = false;
            this.shiftDecRButton.Enabled = false;

            version.Label = "Demo period expired";
        }
        private void openSC_Click(object sender, RibbonControlEventArgs e)
        {
            bool isVis = Globals.ThisAddIn.shortcutsTaskPane.Visible;

            foreach (Microsoft.Office.Tools.CustomTaskPane taskpane in Globals.ThisAddIn.CustomTaskPanes)
            {
                taskpane.Visible = false;
            }

            if (!isVis)
            {
                Globals.ThisAddIn.shortcutsTaskPane.Visible = true;
            }
        }
        private void openForm_Click(object sender, RibbonControlEventArgs e)
        {
            bool isVis = Globals.ThisAddIn.formatsTaskPane.Visible;

            foreach (Microsoft.Office.Tools.CustomTaskPane taskpane in Globals.ThisAddIn.CustomTaskPanes)
            {
                taskpane.Visible = false;
            }

            if (!isVis)
            {
                Globals.ThisAddIn.formatsTaskPane.Visible = true;
            }
        }
        private void openDK_Click(object sender, RibbonControlEventArgs e)
        {
            bool isVis = Globals.ThisAddIn.disabledKeysTaskpane.Visible;

            foreach (Microsoft.Office.Tools.CustomTaskPane taskpane in Globals.ThisAddIn.CustomTaskPanes)
            {
                taskpane.Visible = false;
            }

            if (!isVis)
            {
                Globals.ThisAddIn.disabledKeysTaskpane.Visible = true;
            }
        }
        private void restartAddIn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Settings_Shutdown();
            Globals.ThisAddIn.Settings_Startup();
        }
        private void enableAddIn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Settings_Startup();
        }
        private void disableAddIn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Settings_Shutdown();
        }
        private void genFormButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleGeneralNumberFormats();
        }

        private void currButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CycleCurrency();
        }
        private void fCurrButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CycleForeignCurrency();
        }
        private void perButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CyclePercent();
        }
        private void multButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleMultiple();
        }
        private void BinButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleBinary();
        }
        private void incDecButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.IncreaseDecimal();
        }
        private void decDecButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.DecreaseDecimal();
        }
        private void toggNegButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleCellSign();
        }
        private void shiftDecLButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ShiftDecimalLeft();
        }
        private void shiftDecRButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ShiftDecimalRight();
        }
    }
}