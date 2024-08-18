using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace GES
{
    /// <summary>
    /// Interaction logic for UserControl4.xaml
    /// </summary>
    public partial class CutsForm : UserControl
    {
        Dictionary<string, Dictionary<string, object>> allFunctions;
        Dictionary<string, Dictionary<string, List<string>>> shortcutsData;
        Dictionary<string, Dictionary<string, List<string>>> defCutsData;
        Dictionary<string, List<string>> userData;
        Dictionary<string, string> keysToEnum;
        Dictionary<string, string> enumToKeys;
        public CutsForm()
        {
            InitializeComponent();

            allFunctions = Globals.ThisAddIn.Functions;
            shortcutsData = Globals.ThisAddIn.ShortcutsData;
            defCutsData = Globals.ThisAddIn.DefShortcutsData;
            userData = Globals.ThisAddIn.FunctionData;

            keysToEnum = Globals.ThisAddIn.KeysToEnum;
            enumToKeys = Globals.ThisAddIn.EnumToKeys;

            cutsContent.Children.Clear();

            foreach (string funName in allFunctions.Keys)
            {
                string fun = (string)allFunctions[funName]["fun"];
                string mainKey = enumToKeys[shortcutsData[fun]["mainKey"][0]]; // pulls the enum key from shortcuts list, then converts to keyboard key
                List<string> modKeys = shortcutsData[fun]["modKeys"];
                bool enabled = bool.Parse(shortcutsData[fun]["enabled"][0]);

                if (userData.Keys.Contains(fun) && userData[fun].Count == 0)
                {
                    enabled = false;
                }

                Grid newGrid = new Grid
                {
                    Name = fun + "Grid",
                    Margin = new Thickness(0, 0, 0, 10)
                };

                #region set column widths
                ColumnDefinition column1 = new ColumnDefinition();
                column1.Width = new GridLength(5, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column1);

                ColumnDefinition column2 = new ColumnDefinition();
                column2.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column2);

                ColumnDefinition column3 = new ColumnDefinition();
                column3.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column3);

                ColumnDefinition column4 = new ColumnDefinition();
                column4.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column4);

                ColumnDefinition column5 = new ColumnDefinition();
                column5.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column5);

                ColumnDefinition column6 = new ColumnDefinition();
                column6.Width = new GridLength(1.5, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column6);

                ColumnDefinition column7 = new ColumnDefinition();
                column7.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column7);

                ColumnDefinition column8 = new ColumnDefinition();
                column8.Width = new GridLength(0.5, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column8);

                TextBlock newTB = new TextBlock
                {
                    Name = fun, // + "TB",
                    Text = funName,
                    Style = FindResource("FunctionName") as Style,
                };
                Grid.SetColumn(newTB, 0);
                newGrid.Children.Add(newTB);

                if (modKeys.Contains("Control"))
                {
                    ToggleButton ctrlToggle = new ToggleButton
                    {
                        Name = fun + "Ctrl",
                        Content = "CTRL",
                        IsChecked = true,
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(ctrlToggle, 1);
                    newGrid.Children.Add(ctrlToggle);
                }
                else
                {
                    ToggleButton ctrlToggle = new ToggleButton
                    {
                        Name = fun + "Ctrl",
                        Content = "CTRL",
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(ctrlToggle, 1);
                    newGrid.Children.Add(ctrlToggle);
                };

                if (modKeys.Contains("Shift"))
                {
                    ToggleButton shiftToggle = new ToggleButton
                    {
                        Name = fun + "Shift",
                        Content = "SHIFT",
                        IsChecked = true,
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(shiftToggle, 2);
                    newGrid.Children.Add(shiftToggle);
                }
                else
                {
                    ToggleButton shiftToggle = new ToggleButton
                    {
                        Name = fun + "Shift",
                        Content = "SHIFT",
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(shiftToggle, 2);
                    newGrid.Children.Add(shiftToggle);
                };

                if (modKeys.Contains("Alt"))
                {
                    ToggleButton altToggle = new ToggleButton
                    {
                        Name = fun + "Alt",
                        Content = "ALT",
                        IsChecked = true,
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(altToggle, 3);
                    newGrid.Children.Add(altToggle);
                }
                else
                {
                    ToggleButton altToggle = new ToggleButton
                    {
                        Name = fun + "Alt",
                        Content = "ALT",
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(altToggle, 3);
                    newGrid.Children.Add(altToggle);
                };

                if (modKeys.Contains("Command"))
                {
                    ToggleButton cmdToggle = new ToggleButton
                    {
                        Name = fun + "Cmd",
                        Content = "CMD",
                        IsChecked = true,
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(cmdToggle, 4);
                    newGrid.Children.Add(cmdToggle);
                }
                else
                {
                    ToggleButton cmdToggle = new ToggleButton
                    {
                        Name = fun + "Cmd",
                        Content = "CMD",
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(cmdToggle, 4);
                    newGrid.Children.Add(cmdToggle);
                }

                TextBox key = new TextBox
                {
                    Name = fun + "Key",
                    Text = mainKey,
                    Style = FindResource("KeyTextBox") as Style,
                };
                key.GotFocus += TextBox_GotFocus;
                key.PreviewLostKeyboardFocus += CheckKeys;
                Grid.SetColumn(key, 5);
                newGrid.Children.Add(key);

                Button reset = new Button
                {
                    Name = fun + "Reset",
                    Content = "Reset",
                    Style = FindResource("ResetButtonStyle2") as Style
                };
                reset.Click += resetSingleCut_Click;
                Grid.SetColumn(reset, 6);
                newGrid.Children.Add(reset);

                CheckBox enable = new CheckBox
                {
                    Name = fun + "Enable",
                    IsChecked = enabled,
                    Style = FindResource("EnableSwitchStyle") as Style
                };
                Grid.SetColumn(enable, 8);
                newGrid.Children.Add(enable);
                #endregion

                cutsContent.Children.Add(newGrid);
            }

            Grid ctrlGrid = new Grid
            {
                Name = "CtrlGrid",
                Margin = new Thickness(0, 20, 0, 0)
            };

            Button saveButton = new Button
            {
                Name = "SaveCuts",
                Content = "Save",
                Style = FindResource("CtrlButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 0, 100, 0),
            };
            saveButton.Click += saveCuts_Click;


            Button resetButton = new Button
            {
                Name = "resetAllCuts",
                Content = "Reset All",
                Style = FindResource("CtrlButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(100, 0, 0, 0),
            };
            resetButton.Click += resetAllCuts_Click;

            ctrlGrid.Children.Add(saveButton);
            ctrlGrid.Children.Add(resetButton);

            cutsContent.Children.Add(ctrlGrid);
        }
        private void saveCuts_Click(object sender, RoutedEventArgs e) // fix foreach. Then replace Enabled in resetAll
        {
            foreach (var child in cutsContent.Children)
            {
                if (child is Grid grid && grid.Name != "CtrlGrid")
                {
                    TextBlock funBlock = grid.Children.OfType<TextBlock>().FirstOrDefault();  //t => t.Name.EndsWith("TB"));
                    ToggleButton ctrlButton = grid.Children.OfType<ToggleButton>().FirstOrDefault(t => t.Name.EndsWith("Ctrl"));
                    ToggleButton shiftButton = grid.Children.OfType<ToggleButton>().FirstOrDefault(t => t.Name.EndsWith("Shift"));
                    ToggleButton altButton = grid.Children.OfType<ToggleButton>().FirstOrDefault(t => t.Name.EndsWith("Alt"));
                    ToggleButton cmdButton = grid.Children.OfType<ToggleButton>().FirstOrDefault(t => t.Name.EndsWith("Cmd"));
                    TextBox keyBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("Key"));
                    CheckBox enableBox = grid.Children.OfType<CheckBox>().FirstOrDefault(t => t.Name.EndsWith("Enable"));

                    string fun = funBlock.Name;
                    bool ctrl = ctrlButton.IsChecked == true;
                    bool shift = shiftButton.IsChecked == true;
                    bool alt = altButton.IsChecked == true;
                    bool cmd = cmdButton.IsChecked == true;
                    string key = keysToEnum[keyBox.Text];
                    bool enabledBool = enableBox.IsChecked == true;

                    if (userData.Keys.Contains(fun) && userData[fun].Count == 0) enabledBool = false;

                    List<string> mainKey = new List<string> { key };
                    List<string> modKeys = new List<string>();
                    List<string> enabled = new List<string> { enabledBool.ToString() };

                    if (ctrl) { modKeys.Add("Control"); }
                    if (shift) { modKeys.Add("Shift"); }
                    if (alt) { modKeys.Add("Alt"); }
                    if (cmd) { modKeys.Add("Command"); }

                    shortcutsData[fun]["mainKey"] = mainKey;
                    shortcutsData[fun]["modKeys"] = modKeys;

                    bool isDup = false;
                    foreach (string function in shortcutsData.Keys)
                    {
                        if (function == fun) continue;

                        List<string> funMain = shortcutsData[function]["mainKey"];
                        List<string> funMods = shortcutsData[function]["modKeys"];
                        List<string> funEnabled = shortcutsData[function]["enabled"];

                        if (mainKey.SequenceEqual(funMain) && modKeys.SequenceEqual(funMods) && funEnabled[0] == "True")
                        {
                            isDup = true;
                            break;
                        }
                    }
                    if (!isDup)
                    {
                        shortcutsData[fun]["enabled"] = enabled;
                    }
                    else
                    {
                        shortcutsData[fun]["enabled"] = new List<string> { false.ToString() };
                    }
                }
            }
            Globals.ThisAddIn.shortcutsTaskPane.Visible = false;
            try
            {
                string shortcutsJSON = JsonConvert.SerializeObject(shortcutsData, Formatting.Indented);
                File.WriteAllText(Globals.ThisAddIn.ShortcutsPath, shortcutsJSON);
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ShowErrorMessage(ex.Message);
            }
        }
        private void resetSingleCut_Click(object sender, RoutedEventArgs e)
        {
            Button clickedButton = (Button)sender;
            Grid newGrid = (Grid)clickedButton.Parent; // parent grid. Names newGrid so works with the reused code

            string funName = newGrid.Children.OfType<TextBlock>().FirstOrDefault().Text;
            string fun = newGrid.Name.Replace("Grid", "");

            // delete children of parent, add default settings to parent
            newGrid.Children.Clear();

            string mainKey = enumToKeys[defCutsData[fun]["mainKey"][0]]; // pulls the enum key from shortcuts list, then converts to keyboard key
            List<string> modKeys = defCutsData[fun]["modKeys"];
            bool enabled = bool.Parse(shortcutsData[fun]["enabled"][0]);

            if (userData.Keys.Contains(fun) && userData[fun].Count == 0)
            {
                enabled = false;
            }

            #region create grid
            TextBlock newTB = new TextBlock
            {
                Name = fun, // + "TB",
                Text = funName,
                Style = FindResource("FunctionName") as Style,
            };
            Grid.SetColumn(newTB, 0);
            newGrid.Children.Add(newTB);

            if (modKeys.Contains("Control"))
            {
                ToggleButton ctrlToggle = new ToggleButton
                {
                    Name = fun + "Ctrl",
                    Content = "CTRL",
                    IsChecked = true,
                    Style = FindResource("ToggleButtonStyle2") as Style
                };
                Grid.SetColumn(ctrlToggle, 1);
                newGrid.Children.Add(ctrlToggle);
            }
            else
            {
                ToggleButton ctrlToggle = new ToggleButton
                {
                    Name = fun + "Ctrl",
                    Content = "CTRL",
                    Style = FindResource("ToggleButtonStyle2") as Style
                };
                Grid.SetColumn(ctrlToggle, 1);
                newGrid.Children.Add(ctrlToggle);
            };

            if (modKeys.Contains("Shift"))
            {
                ToggleButton shiftToggle = new ToggleButton
                {
                    Name = fun + "Shift",
                    Content = "SHIFT",
                    IsChecked = true,
                    Style = FindResource("ToggleButtonStyle2") as Style
                };
                Grid.SetColumn(shiftToggle, 2);
                newGrid.Children.Add(shiftToggle);
            }
            else
            {
                ToggleButton shiftToggle = new ToggleButton
                {
                    Name = fun + "Shift",
                    Content = "SHIFT",
                    Style = FindResource("ToggleButtonStyle2") as Style
                };
                Grid.SetColumn(shiftToggle, 2);
                newGrid.Children.Add(shiftToggle);
            };

            if (modKeys.Contains("Alt"))
            {
                ToggleButton altToggle = new ToggleButton
                {
                    Name = fun + "Alt",
                    Content = "ALT",
                    IsChecked = true,
                    Style = FindResource("ToggleButtonStyle2") as Style
                };
                Grid.SetColumn(altToggle, 3);
                newGrid.Children.Add(altToggle);
            }
            else
            {
                ToggleButton altToggle = new ToggleButton
                {
                    Name = fun + "Alt",
                    Content = "ALT",
                    Style = FindResource("ToggleButtonStyle2") as Style
                };
                Grid.SetColumn(altToggle, 3);
                newGrid.Children.Add(altToggle);
            };

            if (modKeys.Contains("Command"))
            {
                ToggleButton cmdToggle = new ToggleButton
                {
                    Name = fun + "Cmd",
                    Content = "CMD",
                    IsChecked = true,
                    Style = FindResource("ToggleButtonStyle2") as Style
                };
                Grid.SetColumn(cmdToggle, 4);
                newGrid.Children.Add(cmdToggle);
            }
            else
            {
                ToggleButton cmdToggle = new ToggleButton
                {
                    Name = fun + "Cmd",
                    Content = "CMD",
                    Style = FindResource("ToggleButtonStyle2") as Style
                };
                Grid.SetColumn(cmdToggle, 4);
                newGrid.Children.Add(cmdToggle);
            }

            TextBox key = new TextBox
            {
                Name = fun + "Key",
                Text = mainKey,
                Style = FindResource("KeyTextBox") as Style,
            };
            key.GotFocus += TextBox_GotFocus;
            key.PreviewLostKeyboardFocus += CheckKeys;
            Grid.SetColumn(key, 5);
            newGrid.Children.Add(key);

            Button reset = new Button
            {
                Name = fun + "Reset",
                Content = "Reset",
                Style = FindResource("ResetButtonStyle2") as Style
            };
            reset.Click += resetSingleCut_Click;
            Grid.SetColumn(reset, 6);
            newGrid.Children.Add(reset);

            CheckBox enable = new CheckBox
            {
                Name = fun + "Enable",
                IsChecked = enabled,
                Style = FindResource("EnableSwitchStyle") as Style
            };
            Grid.SetColumn(enable, 8);
            newGrid.Children.Add(enable);

            #endregion
        }
        private void resetAllCuts_Click(object sender, RoutedEventArgs e)
        {
            cutsContent.Children.Clear();

            foreach (string funName in allFunctions.Keys)
            {
                string fun = (string)allFunctions[funName]["fun"];
                string mainKey = enumToKeys[defCutsData[fun]["mainKey"][0]]; // pulls the enum key from shortcuts list, then converts to keyboard key
                List<string> modKeys = defCutsData[fun]["modKeys"];
                //bool enabled = bool.Parse(defCutsData[fun]["enabled"][0]);
                bool enabled = bool.Parse(shortcutsData[fun]["enabled"][0]);

                if (userData.Keys.Contains(fun) && userData[fun].Count == 0)
                {
                    enabled = false;
                }

                Grid newGrid = new Grid
                {
                    Margin = new Thickness(0, 0, 0, 10)
                };

                #region set column widths
                ColumnDefinition column1 = new ColumnDefinition();
                column1.Width = new GridLength(5, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column1);

                ColumnDefinition column2 = new ColumnDefinition();
                column2.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column2);

                ColumnDefinition column3 = new ColumnDefinition();
                column3.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column3);

                ColumnDefinition column4 = new ColumnDefinition();
                column4.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column4);

                ColumnDefinition column5 = new ColumnDefinition();
                column5.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column5);

                ColumnDefinition column6 = new ColumnDefinition();
                column6.Width = new GridLength(1.5, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column6);

                ColumnDefinition column7 = new ColumnDefinition();
                column7.Width = new GridLength(1, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column7);

                ColumnDefinition column8 = new ColumnDefinition();
                column8.Width = new GridLength(0.5, GridUnitType.Star);
                newGrid.ColumnDefinitions.Add(column8);

                TextBlock newTB = new TextBlock
                {
                    Name = fun, // + "TB",
                    Text = funName,
                    Style = FindResource("FunctionName") as Style,
                };
                Grid.SetColumn(newTB, 0);
                newGrid.Children.Add(newTB);

                if (modKeys.Contains("Control"))
                {
                    ToggleButton ctrlToggle = new ToggleButton
                    {
                        Name = fun + "Ctrl",
                        Content = "CTRL",
                        IsChecked = true,
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(ctrlToggle, 1);
                    newGrid.Children.Add(ctrlToggle);
                }
                else
                {
                    ToggleButton ctrlToggle = new ToggleButton
                    {
                        Name = fun + "Ctrl",
                        Content = "CTRL",
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(ctrlToggle, 1);
                    newGrid.Children.Add(ctrlToggle);
                };

                if (modKeys.Contains("Shift"))
                {
                    ToggleButton shiftToggle = new ToggleButton
                    {
                        Name = fun + "Shift",
                        Content = "SHIFT",
                        IsChecked = true,
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(shiftToggle, 2);
                    newGrid.Children.Add(shiftToggle);
                }
                else
                {
                    ToggleButton shiftToggle = new ToggleButton
                    {
                        Name = fun + "Shift",
                        Content = "SHIFT",
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(shiftToggle, 2);
                    newGrid.Children.Add(shiftToggle);
                };

                if (modKeys.Contains("Alt"))
                {
                    ToggleButton altToggle = new ToggleButton
                    {
                        Name = fun + "Alt",
                        Content = "ALT",
                        IsChecked = true,
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(altToggle, 3);
                    newGrid.Children.Add(altToggle);
                }
                else
                {
                    ToggleButton altToggle = new ToggleButton
                    {
                        Name = fun + "Alt",
                        Content = "ALT",
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(altToggle, 3);
                    newGrid.Children.Add(altToggle);
                };

                if (modKeys.Contains("Command"))
                {
                    ToggleButton cmdToggle = new ToggleButton
                    {
                        Name = fun + "Cmd",
                        Content = "CMD",
                        IsChecked = true,
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(cmdToggle, 4);
                    newGrid.Children.Add(cmdToggle);
                }
                else
                {
                    ToggleButton cmdToggle = new ToggleButton
                    {
                        Name = fun + "Cmd",
                        Content = "CMD",
                        Style = FindResource("ToggleButtonStyle2") as Style
                    };
                    Grid.SetColumn(cmdToggle, 4);
                    newGrid.Children.Add(cmdToggle);
                }

                TextBox key = new TextBox
                {
                    Name = fun + "Key",
                    Text = mainKey,
                    Style = FindResource("KeyTextBox") as Style,
                };
                key.GotFocus += TextBox_GotFocus;
                key.PreviewLostKeyboardFocus += CheckKeys;
                Grid.SetColumn(key, 5);
                newGrid.Children.Add(key);

                Button reset = new Button
                {
                    Name = fun + "Reset",
                    Content = "Reset",
                    Style = FindResource("ResetButtonStyle2") as Style
                };
                reset.Click += resetSingleCut_Click;
                Grid.SetColumn(reset, 6);
                newGrid.Children.Add(reset);

                CheckBox enable = new CheckBox
                {
                    Name = fun + "Enable",
                    IsChecked = enabled,
                    Style = FindResource("EnableSwitchStyle") as Style
                };
                Grid.SetColumn(enable, 8);
                newGrid.Children.Add(enable);
                #endregion

                cutsContent.Children.Add(newGrid);
            }

            Grid ctrlGrid = new Grid
            {
                Name = "CtrlGrid",
                Margin = new Thickness(0, 20, 0, 0)
            };

            Button saveButton = new Button
            {
                Name = "SaveCuts",
                Content = "Save",
                Style = FindResource("CtrlButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 0, 100, 0),
            };
            saveButton.Click += saveCuts_Click;

            Button resetButton = new Button
            {
                Name = "resetAllCuts",
                Content = "Reset All",
                Style = FindResource("CtrlButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(100, 0, 0, 0),
            };
            resetButton.Click += resetAllCuts_Click;

            ctrlGrid.Children.Add(saveButton);
            ctrlGrid.Children.Add(resetButton);

            cutsContent.Children.Add(ctrlGrid);
        }
        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
        }
        private void CheckKeys(object sender, RoutedEventArgs e)
        {
            TextBox tb = (TextBox)sender;
            string key = tb.Text;

            // Check if key is in dictionary
            if (!keysToEnum.ContainsKey(key))
            {
                // Indicate the error visually
                tb.BorderBrush = new SolidColorBrush(Colors.Red);
                tb.Opacity = .6;

                // Refocus and return
                tb.Focus();
                e.Handled = true;
            }
        }
    }
}
