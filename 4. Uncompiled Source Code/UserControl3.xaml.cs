using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace GES
{
    public partial class FormatForm : UserControl
    {
        public string selectedItem;
        List<string> numFunctions = new List<string>();
        List<string> colFunctions = new List<string>();
        List<string> othNumFunctions = new List<string>();
        List<string> othColFunctions = new List<string>();
        List<string> autocolorFunctions = new List<string>();
        Dictionary<string, Dictionary<string, List<string>>> shortcutsData;
        Dictionary<string, Dictionary<string, object>> allFunctions;
        private FormatFormPopup formatFormPopup;

        public FormatForm()
        {
            InitializeComponent();

            allFunctions = Globals.ThisAddIn.Functions;
            shortcutsData = Globals.ThisAddIn.ShortcutsData;

            foreach (string key in allFunctions.Keys)
            {
                if ((string)allFunctions[key]["type"] == "number")
                {
                    numFunctions.Add(key);
                }
                else if ((string)allFunctions[key]["type"] == "color")
                {
                    colFunctions.Add(key);
                }
                else if ((string)allFunctions[key]["type"] == "otherNum")
                {
                    othNumFunctions.Add(key);
                }
                else if ((string)allFunctions[key]["type"] == "otherCol")
                {
                    othColFunctions.Add(key);
                }
                else if ((string)allFunctions[key]["type"] == "autocolor")
                {
                    autocolorFunctions.Add(key);
                }

                if (numFunctions.Contains(key) || colFunctions.Contains(key) || othNumFunctions.Contains(key) || othColFunctions.Contains(key) || autocolorFunctions.Contains(key))
                {
                    ListBoxItem newLBI = new ListBoxItem
                    {
                        Content = key,
                        Style = FindResource("FunctionItem") as Style,
                    };
                    FunctionList.Items.Add(newLBI);
                }
            }
        }
        private async void FunctionList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FunctionList.SelectedItem != null)
            {
                // Allow the UI time to refresh
                await Task.Delay(10);

                selectedItem = (FunctionList.SelectedItem as ListBoxItem).Content.ToString();
                formatFormPopup = ShowFormatWindow(selectedItem, this);
            }
        }
        public FormatFormPopup ShowFormatWindow(string data, FormatForm formatForm)
        {
            formatFormPopup = new FormatFormPopup(data, formatForm);

            // Set the owner of the WPF window to the Excel application.
            WindowInteropHelper windowInteropHelper = new WindowInteropHelper(formatFormPopup);
            windowInteropHelper.Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd);

            formatFormPopup.Closed += FormatFormPopup_Closed;
            formatFormPopup.ShowDialog();

            return formatFormPopup;
        }
        private void FormatFormPopup_Closed(object sender, EventArgs e)
        {
            FunctionList.SelectedItem = null;
        }
    }
}