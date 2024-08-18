using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using Xceed.Wpf.Toolkit.Primitives;

namespace GES
{
    /// <summary>
    /// Interaction logic for UserControl7.xaml
    /// </summary>
    public partial class DisabledKeysForm : System.Windows.Controls.UserControl
    {
        public Dictionary<string, string> keysToEnum;
        public Dictionary<string, string> enumToKeys;
        public List<string> disabledKeys;
        public DisabledKeysForm()
        {
            InitializeComponent();

            keysToEnum = Globals.ThisAddIn.KeysToEnum;
            enumToKeys = Globals.ThisAddIn.EnumToKeys;
            disabledKeys = Globals.ThisAddIn.DisabledKeys["keys"];

            keysContent.Children.Clear();

            Label label = new Label
            {
                Content = "Disabled Keys",
                Style = FindResource("ContentTitle") as Style,

            };
            keysContent.Children.Add(label);

            for (int gridCount = 0; gridCount < disabledKeys.Count; gridCount++)
            {
                Grid newGrid = new Grid
                {
                    Name = "Grid" + (gridCount + 1).ToString()
                };

                TextBox newTB = new TextBox
                {
                    Name = "TB" + (gridCount + 1).ToString(),
                    Text = disabledKeys[gridCount],
                    Style = FindResource("KeyTextBox") as Style,
                    Margin = new Thickness(0, 5, 0, 5)
                };
                newTB.PreviewLostKeyboardFocus += CheckKeys;

                Button deleteButton = new Button
                {
                    Name = "DB" + (gridCount + 1).ToString(),
                    Content = "X",
                    Style = FindResource("RemoveButtonStyle") as Style,
                    Margin = new Thickness(100, 0, 0, 0)
                };

                deleteButton.Click += RemoveButton_Click;

                newGrid.Children.Add(newTB);
                newGrid.Children.Add(deleteButton);

                keysContent.Children.Add(newGrid);
            }

            Grid ctrlGrid = new Grid
            {
                Margin = new Thickness(0, 10, 0, 0)
            };

            Button addButton = new Button
            {
                Name = "addKey",
                Content = "+",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 0, 0, 30)
            };
            addButton.Click += addKey_Click;

            Button saveButton = new Button
            {
                Name = "saveKey",
                Content = "Save",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 40, 0, 0)
            };
            saveButton.Click += saveKeys_Click;

            ctrlGrid.Children.Add(addButton);
            ctrlGrid.Children.Add(saveButton);

            keysContent.Children.Add(ctrlGrid);
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
        public void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            Button clickedButton = (Button)sender;
            Grid parentGrid = (Grid)clickedButton.Parent;
            keysContent.Children.Remove(parentGrid);
        }
        private void addKey_Click(object sender, RoutedEventArgs e)
        {
            int gridCount = 0;
            foreach (var child in keysContent.Children)
            {
                if (child is Grid)
                {
                    gridCount++;
                }
            }

            Grid newGrid = new Grid
            {
                Name = "Grid" + (gridCount + 1).ToString()
            };

            TextBox newTB = new TextBox
            {
                Name = "TB" + (gridCount + 1).ToString(),
                Text = "",
                Style = FindResource("KeyTextBox") as Style,
                Margin = new Thickness(0, 5, 0, 5)
            };
            newTB.PreviewLostKeyboardFocus += CheckKeys;

            Button deleteButton = new Button
            {
                Name = "DB" + (gridCount + 1).ToString(),
                Content = "X",
                Style = FindResource("RemoveButtonStyle") as Style,
                Margin = new Thickness(100, 0, 0, 0)
            };

            deleteButton.Click += RemoveButton_Click;

            newGrid.Children.Add(newTB);
            newGrid.Children.Add(deleteButton);

            keysContent.Children.Insert(keysContent.Children.Count - 1, newGrid);
        }
        private void saveKeys_Click(object sender, RoutedEventArgs e)
        {
            List<String> newKeys = new List<String>();

            foreach (var child in keysContent.Children)
            {
                if (child is Grid grid)
                {
                    foreach (var child2 in grid.Children)
                    {
                        if (child2 is TextBox textBox)
                        {
                            if (textBox.Text != "")
                            {
                                newKeys.Add(textBox.Text);
                                //textBox.Text = "";
                            }
                        }
                    }
                }
            }
            saveFun(newKeys);
        }
        public void saveFun(List<string> data)
        {
            Dictionary<string, List<string>> disabledKeysDict = new Dictionary<string, List<string>>
            {
                { "keys", data }
            };
            try
            {
                string keysJson = JsonConvert.SerializeObject(disabledKeysDict, Formatting.Indented);
                File.WriteAllText(System.IO.Path.Combine(Globals.ThisAddIn.appData, "disabledKeys.json"), keysJson);
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ShowErrorMessage(ex.Message);
            }

            Globals.ThisAddIn.disabledKeysTaskpane.Visible = false;
        }
    }
}
