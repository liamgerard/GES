using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace GES
{
    /// <summary>
    /// Interaction logic for UserControl5.xaml
    /// </summary>
    public partial class FormatFormPopup : System.Windows.Window
    {
        string selectedItem;
        List<string> numFunctions = new List<string>();
        List<string> colFunctions = new List<string>();
        List<string> othNumFunctions = new List<string>();
        List<string> othColFunctions = new List<string>();
        List<string> autocolorFunctions = new List<string>();
        Dictionary<string, Dictionary<string, object>> allFunctions;
        Dictionary<string, Dictionary<string, List<string>>> shortcutsData;
        Dictionary<string, List<string>> userData;
        FormatForm callerForm;
        private ColorFormPopup colorFormPopup;
        public FormatFormPopup(string selectedItemPar, FormatForm callerFormPar)
        {
            InitializeComponent();

            allFunctions = Globals.ThisAddIn.Functions;
            shortcutsData = Globals.ThisAddIn.ShortcutsData;
            userData = Globals.ThisAddIn.FunctionData;

            selectedItem = selectedItemPar;
            callerForm = callerFormPar;

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
            }

            if (numFunctions.Contains(selectedItem))
            {
                numContent.Children.Clear();

                Label label = new Label
                {
                    Content = selectedItem,
                    Style = FindResource("ContentTitle") as Style,

                };
                numContent.Children.Add(label);

                string fun = (string)allFunctions[selectedItem]["fun"];
                int countForm = userData[fun].Count;

                for (int gridCount = 0; gridCount < countForm; gridCount++)
                {
                    Grid newGrid = new Grid
                    {
                        Name = "Grid" + (gridCount + 1).ToString(),
                    };

                    TextBox newTB = new TextBox
                    {
                        Name = "TB" + (gridCount + 1).ToString(),
                        Text = userData[fun][gridCount],
                        Style = FindResource("NumFormTextBox") as Style
                    };

                    Button deleteButton = new Button
                    {
                        Name = "DB" + (gridCount + 1).ToString(),
                        Content = "X",
                        Style = FindResource("RemoveButtonStyle") as Style,
                    };

                    deleteButton.Click += RemoveButton_Click;

                    newGrid.Children.Add(newTB);
                    newGrid.Children.Add(deleteButton);

                    numContent.Children.Add(newGrid);
                }

                Button addButton = new Button
                {
                    Name = "addNum",
                    Content = "+",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                };
                addButton.Click += addNum_Click;

                Grid ctrlGrid = new Grid
                {
                    Margin = new Thickness(0, 20, 0, 0)
                };

                Button saveButton = new Button
                {
                    Name = "saveNum",
                    Content = "Save",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(0, 0, 100, 0),
                };
                saveButton.Click += saveNum_Click;

                Button resetButton = new Button
                {
                    Name = "resetNum",
                    Content = "Reset",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(100, 0, 0, 0),
                };
                resetButton.Click += resetNum_Click;

                ctrlGrid.Children.Add(saveButton);
                ctrlGrid.Children.Add(resetButton);

                numContent.Children.Add(addButton);
                numContent.Children.Add(ctrlGrid);

                // make content visible
                numContent.Visibility = Visibility.Visible;
                colContent.Visibility = Visibility.Collapsed;
                othContent.Visibility = Visibility.Collapsed;
            }
            else if (colFunctions.Contains(selectedItem))
            {
                colContent.Children.Clear();

                Label label = new Label
                {
                    Content = selectedItem,
                    Style = FindResource("ContentTitle") as Style,

                };
                colContent.Children.Add(label);

                // based on the selected item, identify the function name
                string fun = (string)allFunctions[selectedItem]["fun"];

                // iterate over data stored in function and add each color to the appropriate textboxes
                int countForm = userData[fun].Count;

                for (int gridCount = 0; gridCount < countForm; gridCount++)
                {
                    Grid newGrid = new Grid
                    {
                        Name = "c" + (gridCount + 1).ToString(),
                    };

                    int rgb = int.Parse(userData[fun][gridCount]);

                    int b = (rgb >> 16) & 255;
                    int g = (rgb >> 8) & 255;
                    int r = rgb & 255;

                    Button colDis = new Button
                    {
                        Style = FindResource("ColorDisplay") as Style,
                    };
                    colDis.Click += colButton_Click;

                    TextBox newTB1 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "R", r.ToString(), new Thickness(0, 5, 110, 5));
                    TextBox newTB2 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "G", g.ToString(), new Thickness(0, 5, 0, 5));
                    TextBox newTB3 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "B", b.ToString(), new Thickness(110, 5, 0, 5));

                    Button deleteButton = new Button
                    {
                        Name = "DB" + (gridCount + 1).ToString(),
                        Content = "X",
                        Style = FindResource("RemoveButtonStyle") as Style,
                    };
                    deleteButton.Click += RemoveButton_Click;

                    newGrid.Children.Add(colDis);
                    newGrid.Children.Add(newTB1);
                    newGrid.Children.Add(newTB2);
                    newGrid.Children.Add(newTB3);
                    newGrid.Children.Add(deleteButton);

                    colContent.Children.Add(newGrid);
                    RGBTextBox(newGrid);
                }

                Button addButton = new Button
                {
                    Name = "addCol",
                    Content = "+",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                };
                addButton.Click += addCol_Click;

                Grid ctrlGrid = new Grid
                {
                    Margin = new Thickness(0, 20, 0, 0)
                };

                Button saveButton = new Button
                {
                    Name = "saveCol",
                    Content = "Save",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(0, 0, 100, 0),
                };
                saveButton.Click += saveCol_Click;

                Button resetButton = new Button
                {
                    Name = "resetCol",
                    Content = "Reset",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(100, 0, 0, 0),
                };
                resetButton.Click += resetCol_Click;

                ctrlGrid.Children.Add(saveButton);
                ctrlGrid.Children.Add(resetButton);

                colContent.Children.Add(addButton);
                colContent.Children.Add(ctrlGrid);

                numContent.Visibility = Visibility.Collapsed;
                colContent.Visibility = Visibility.Visible;
                othContent.Visibility = Visibility.Collapsed;
            }
            else if (othNumFunctions.Contains(selectedItem))
            {
                othContent.Children.Clear();

                Label label = new Label
                {
                    Content = selectedItem,
                    Style = FindResource("ContentTitle") as Style,

                };
                othContent.Children.Add(label);

                string fun = (string)allFunctions[selectedItem]["fun"];
                int countForm = userData[fun].Count;

                for (int gridCount = 0; gridCount < countForm; gridCount++)
                {
                    Grid newGrid = new Grid
                    {
                        Name = "Grid" + (gridCount + 1).ToString(),
                    };

                    TextBox newTB = new TextBox
                    {
                        Name = "TB" + (gridCount + 1).ToString(),
                        Text = userData[fun][gridCount],
                        Style = FindResource("NumFormTextBox") as Style
                    };

                    newGrid.Children.Add(newTB);

                    othContent.Children.Add(newGrid);
                }

                Grid ctrlGrid = new Grid
                {
                    Margin = new Thickness(0, 20, 0, 0)
                };

                Button saveButton = new Button
                {
                    Name = "saveOthNum",
                    Content = "Save",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(0, 0, 100, 0),
                };
                saveButton.Click += saveOthNum_Click;

                Button resetButton = new Button
                {
                    Name = "resetOthNum",
                    Content = "Reset",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(100, 0, 0, 0),
                };
                resetButton.Click += resetOthNum_Click;

                ctrlGrid.Children.Add(saveButton);
                ctrlGrid.Children.Add(resetButton);

                othContent.Children.Add(ctrlGrid);

                // make content visible
                numContent.Visibility = Visibility.Collapsed;
                colContent.Visibility = Visibility.Collapsed;
                othContent.Visibility = Visibility.Visible;
            }
            else if (othColFunctions.Contains(selectedItem))
            {
                othContent.Children.Clear();

                Label label = new Label
                {
                    Content = selectedItem,
                    Style = FindResource("ContentTitle") as Style,

                };
                othContent.Children.Add(label);

                // based on the selected item, identify the function name
                string fun = (string)allFunctions[selectedItem]["fun"];

                // iterate over data stored in function and add each color to the appropriate textboxes
                int countForm = userData[fun].Count;

                for (int gridCount = 0; gridCount < countForm; gridCount++)
                {
                    Grid newGrid = new Grid
                    {
                        Name = "c" + (gridCount + 1).ToString(),
                    };

                    int rgb = int.Parse(userData[fun][gridCount]);

                    int b = (rgb >> 16) & 255;
                    int g = (rgb >> 8) & 255;
                    int r = rgb & 255;

                    Button colDis = new Button
                    {
                        Style = FindResource("ColorDisplay") as Style,
                    };
                    colDis.Click += colButton_Click;

                    TextBox newTB1 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "R", r.ToString(), new Thickness(0, 5, 110, 5));
                    TextBox newTB2 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "G", g.ToString(), new Thickness(0, 5, 0, 5));
                    TextBox newTB3 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "B", b.ToString(), new Thickness(110, 5, 0, 5));

                    newGrid.Children.Add(colDis);
                    newGrid.Children.Add(newTB1);
                    newGrid.Children.Add(newTB2);
                    newGrid.Children.Add(newTB3);

                    othContent.Children.Add(newGrid);
                    RGBTextBox(newGrid);
                }

                Grid ctrlGrid = new Grid
                {
                    Margin = new Thickness(0, 20, 0, 0)
                };

                Button saveButton = new Button
                {
                    Name = "saveOthCol",
                    Content = "Save",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(0, 0, 100, 0),
                };
                saveButton.Click += saveOthCol_Click;

                Button resetButton = new Button
                {
                    Name = "resetOthCol",
                    Content = "Reset",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(100, 0, 0, 0),
                };
                resetButton.Click += resetOthCol_Click;

                ctrlGrid.Children.Add(saveButton);
                ctrlGrid.Children.Add(resetButton);

                othContent.Children.Add(ctrlGrid);

                numContent.Visibility = Visibility.Collapsed;
                colContent.Visibility = Visibility.Collapsed;
                othContent.Visibility = Visibility.Visible;
            }
            else if (autocolorFunctions.Contains(selectedItem))
            {
                othContent.Children.Clear();

                Label label = new Label
                {
                    Content = selectedItem,
                    Style = FindResource("ContentTitle") as Style,

                };
                othContent.Children.Add(label);

                // based on the selected item, identify the function name
                string fun = (string)allFunctions[selectedItem]["fun"];

                // iterate over data stored in function and add each color to the appropriate textboxes
                int countForm = userData[fun].Count;
                List<string> labels = new List<string> { "Hardcode", "Partial Input", "Formula", "Sheet Reference", "File Reference" };

                for (int gridCount = 0; gridCount < countForm; gridCount++)
                {
                    Grid newGrid = new Grid
                    {
                        Name = "c" + (gridCount + 1).ToString(),
                    };

                    int rgb = int.Parse(userData[fun][gridCount]);

                    int b = (rgb >> 16) & 255;
                    int g = (rgb >> 8) & 255;
                    int r = rgb & 255;

                    Button colDis = new Button
                    {
                        Style = FindResource("ColorDisplay") as Style,
                    };
                    colDis.Click += colButton_Click;

                    TextBox newTB1 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "R", r.ToString(), new Thickness(0, 5, 110, 5));
                    TextBox newTB2 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "G", g.ToString(), new Thickness(0, 5, 0, 5));
                    TextBox newTB3 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "B", b.ToString(), new Thickness(110, 5, 0, 5));

                    TextBlock colLabel = new TextBlock
                    {
                        Style = FindResource("FunctionName") as Style,
                        Margin = new Thickness(290, 5, 0, 5),
                        Text = labels[gridCount],
                        HorizontalAlignment = HorizontalAlignment.Left
                    };

                    newGrid.Children.Add(colDis);
                    newGrid.Children.Add(newTB1);
                    newGrid.Children.Add(newTB2);
                    newGrid.Children.Add(newTB3);
                    newGrid.Children.Add(colLabel);

                    othContent.Children.Add(newGrid);
                    RGBTextBox(newGrid);
                }

                Grid ctrlGrid = new Grid
                {
                    Margin = new Thickness(0, 20, 0, 0)
                };

                Button saveButton = new Button
                {
                    Name = "saveOthCol",
                    Content = "Save",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(0, 0, 100, 0),
                };
                saveButton.Click += saveOthCol_Click;

                Button resetButton = new Button
                {
                    Name = "resetOthCol",
                    Content = "Reset",
                    Style = FindResource("AddButtonStyle") as Style,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Width = 80,
                    Margin = new Thickness(100, 0, 0, 0),
                };
                resetButton.Click += resetAutoCol_Click;

                ctrlGrid.Children.Add(saveButton);
                ctrlGrid.Children.Add(resetButton);

                othContent.Children.Add(ctrlGrid);

                numContent.Visibility = Visibility.Collapsed;
                colContent.Visibility = Visibility.Collapsed;
                othContent.Visibility = Visibility.Visible;
            }

        }
        private void addNum_Click(object sender, RoutedEventArgs e)
        {
            int gridCount = 0;
            foreach (var child in numContent.Children)
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
                Style = FindResource("NumFormTextBox") as Style
            };

            Button deleteButton = new Button
            {
                Name = "DB" + (gridCount + 1).ToString(),
                Content = "X",
                Style = FindResource("RemoveButtonStyle") as Style,
            };

            deleteButton.Click += RemoveButton_Click;

            newGrid.Children.Add(newTB);
            newGrid.Children.Add(deleteButton);

            numContent.Children.Insert(numContent.Children.Count - 2, newGrid);
        }
        private void saveNum_Click(object sender, RoutedEventArgs e)
        {
            List<String> newForms = new List<String>();

            foreach (var child in numContent.Children)
            {
                if (child is Grid grid)
                {
                    foreach (var child2 in grid.Children)
                    {
                        if (child2 is TextBox textBox)
                        {
                            if (textBox.Text != "")
                            {
                                newForms.Add(textBox.Text);
                                textBox.Text = "";
                            }
                        }
                    }
                }
            }

            saveFun(newForms);
        }
        private void resetNum_Click(object sender, RoutedEventArgs e)
        {
            numContent.Children.Clear();

            Label label = new Label
            {
                Content = selectedItem,
                Style = FindResource("ContentTitle") as Style,

            };
            numContent.Children.Add(label);

            //string jsonPath = System.IO.Path.Combine(baseDirectory, "defaultData.json");
            //string json = File.ReadAllText(jsonPath);
            //Dictionary<string, List<string>> defaultData = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);

            Dictionary<string, List<string>> defaultData = Globals.ThisAddIn.DefFunctionData;

            // clear and remake the content page
            string fun = (string)allFunctions[selectedItem]["fun"];
            int countForm = defaultData[fun].Count;

            for (int gridCount = 0; gridCount < countForm; gridCount++)
            {
                Grid newGrid = new Grid
                {
                    Name = "Grid" + (gridCount + 1).ToString(),
                };

                TextBox newTB = new TextBox
                {
                    Name = "TB" + (gridCount + 1).ToString(),
                    Text = defaultData[fun][gridCount],
                    Style = FindResource("NumFormTextBox") as Style
                };

                Button deleteButton = new Button
                {
                    Name = "DB" + (gridCount + 1).ToString(),
                    Content = "X",
                    Style = FindResource("RemoveButtonStyle") as Style,
                };

                deleteButton.Click += RemoveButton_Click;

                newGrid.Children.Add(newTB);
                newGrid.Children.Add(deleteButton);

                numContent.Children.Insert(numContent.Children.Count, newGrid);
            }

            Button addButton = new Button
            {
                Name = "addNum",
                Content = "+",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
            };
            addButton.Click += addNum_Click;

            Grid ctrlGrid = new Grid
            {
                Margin = new Thickness(0, 20, 0, 0)
            };

            Button saveButton = new Button
            {
                Name = "saveNum",
                Content = "Save",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 0, 100, 0),
            };

            saveButton.Click += saveNum_Click;

            Button resetButton = new Button
            {
                Name = "resetNum",
                Content = "Reset",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(100, 0, 0, 0),
            };

            resetButton.Click += resetNum_Click;

            ctrlGrid.Children.Add(saveButton);
            ctrlGrid.Children.Add(resetButton);

            numContent.Children.Add(addButton);
            numContent.Children.Add(ctrlGrid);
        }
        private void addCol_Click(object sender, RoutedEventArgs e)
        {
            int gridCount = 0;
            foreach (var child in colContent.Children)
            {
                if (child is Grid)
                {
                    gridCount++;
                }
            }

            Grid newGrid = new Grid
            {
                Name = "c" + (gridCount + 1).ToString(),
            };

            Button colDis = new Button
            {
                Style = FindResource("ColorDisplay") as Style,
            };
            colDis.Click += colButton_Click;

            TextBox newTB1 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "R", "", new Thickness(0, 5, 110, 5));
            TextBox newTB2 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "G", "", new Thickness(0, 5, 0, 5));
            TextBox newTB3 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "B", "", new Thickness(110, 5, 0, 5));

            Button deleteButton = new Button
            {
                Name = "DB" + (gridCount + 1).ToString(),
                Content = "X",
                Style = FindResource("RemoveButtonStyle") as Style,
            };

            deleteButton.Click += RemoveButton_Click;

            newGrid.Children.Add(colDis);
            newGrid.Children.Add(newTB1);
            newGrid.Children.Add(newTB2);
            newGrid.Children.Add(newTB3);
            newGrid.Children.Add(deleteButton);

            colContent.Children.Insert(colContent.Children.Count - 2, newGrid);
        }
        private void saveCol_Click(object sender, RoutedEventArgs e)
        {
            List<string> newCols = new List<string>();

            foreach (var child in colContent.Children)
            {
                if (child is Grid grid)
                {
                    TextBox rTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("R"));
                    TextBox gTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("G"));
                    TextBox bTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("B"));

                    if (rTextBox != null && gTextBox != null && bTextBox != null)
                    {
                        if (int.TryParse(rTextBox.Text, out int r) && int.TryParse(gTextBox.Text, out int g) && int.TryParse(bTextBox.Text, out int b))
                        {
                            int rgb = (b << 16) | (g << 8) | r;
                            newCols.Add(rgb.ToString());

                            rTextBox.Text = "";
                            gTextBox.Text = "";
                            bTextBox.Text = "";
                        }
                    }
                }
            }
            saveFun(newCols);
        }
        private void resetCol_Click(object sender, RoutedEventArgs e)
        {
            colContent.Children.Clear();

            Label label = new Label
            {
                Content = selectedItem,
                Style = FindResource("ContentTitle") as Style,

            };
            colContent.Children.Add(label);

            Dictionary<string, List<string>> defaultData = Globals.ThisAddIn.DefFunctionData;

            // clear and remake the content page
            string fun = (string)allFunctions[selectedItem]["fun"];

            // iterate over data stored in function and add each color to the appropriate textboxes
            int countForm = defaultData[fun].Count;

            for (int gridCount = 0; gridCount < countForm; gridCount++)
            {
                Grid newGrid = new Grid
                {
                    Name = "c" + (gridCount + 1).ToString(),
                };

                int rgb = int.Parse(defaultData[fun][gridCount]);

                int b = (rgb >> 16) & 255;
                int g = (rgb >> 8) & 255;
                int r = rgb & 255;

                Button colDis = new Button
                {
                    Style = FindResource("ColorDisplay") as Style,
                };
                colDis.Click += colButton_Click;

                TextBox newTB1 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "R", r.ToString(), new Thickness(0, 5, 110, 5));
                TextBox newTB2 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "G", g.ToString(), new Thickness(0, 5, 0, 5));
                TextBox newTB3 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "B", b.ToString(), new Thickness(110, 5, 0, 5));

                Button deleteButton = new Button
                {
                    Name = "DB" + (gridCount + 1).ToString(),
                    Content = "X",
                    Style = FindResource("RemoveButtonStyle") as Style,
                };

                deleteButton.Click += RemoveButton_Click;

                newGrid.Children.Add(colDis);
                newGrid.Children.Add(newTB1);
                newGrid.Children.Add(newTB2);
                newGrid.Children.Add(newTB3);
                newGrid.Children.Add(deleteButton);

                colContent.Children.Add(newGrid);
                RGBTextBox(newGrid);
            }

            Button addButton = new Button
            {
                Name = "addCol",
                Content = "+",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
            };
            addButton.Click += addCol_Click;

            Grid ctrlGrid = new Grid
            {
                Margin = new Thickness(0, 20, 0, 0)
            };

            Button saveButton = new Button
            {
                Name = "saveCol",
                Content = "Save",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 0, 100, 0),
            };
            saveButton.Click += saveCol_Click;

            Button resetButton = new Button
            {
                Name = "resetCol",
                Content = "Reset",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(100, 0, 0, 0),
            };
            resetButton.Click += resetCol_Click;

            ctrlGrid.Children.Add(saveButton);
            ctrlGrid.Children.Add(resetButton);

            colContent.Children.Add(addButton);
            colContent.Children.Add(ctrlGrid);
        }
        private void saveOthNum_Click(object sender, RoutedEventArgs e)
        {
            List<String> newForms = new List<String>();

            foreach (var child in othContent.Children)
            {
                if (child is Grid grid)
                {
                    foreach (var child2 in grid.Children)
                    {
                        if (child2 is TextBox textBox)
                        {
                            if (textBox.Text != "")
                            {
                                newForms.Add(textBox.Text);
                                textBox.Text = "";
                            }
                        }
                    }
                }
            }

            saveFun(newForms);
        }
        private void resetOthNum_Click(object sender, RoutedEventArgs e)
        {
            othContent.Children.Clear();

            Label label = new Label
            {
                Content = selectedItem,
                Style = FindResource("ContentTitle") as Style,

            };
            othContent.Children.Add(label);

            Dictionary<string, List<string>> defaultData = Globals.ThisAddIn.DefFunctionData;

            // clear and remake the content page
            string fun = (string)allFunctions[selectedItem]["fun"];
            int countForm = defaultData[fun].Count;

            for (int gridCount = 0; gridCount < countForm; gridCount++)
            {
                Grid newGrid = new Grid
                {
                    Name = "Grid" + (gridCount + 1).ToString(),
                };

                TextBox newTB = new TextBox
                {
                    Name = "TB" + (gridCount + 1).ToString(),
                    Text = defaultData[fun][gridCount],
                    Style = FindResource("NumFormTextBox") as Style
                };

                newGrid.Children.Add(newTB);

                othContent.Children.Insert(othContent.Children.Count, newGrid);
            }

            Grid ctrlGrid = new Grid
            {
                Margin = new Thickness(0, 20, 0, 0)
            };

            Button saveButton = new Button
            {
                Name = "saveOthNum",
                Content = "Save",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 0, 100, 0),
            };

            saveButton.Click += saveOthNum_Click;

            Button resetButton = new Button
            {
                Name = "resetOthNum",
                Content = "Reset",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(100, 0, 0, 0),
            };

            resetButton.Click += resetOthNum_Click;

            ctrlGrid.Children.Add(saveButton);
            ctrlGrid.Children.Add(resetButton);

            othContent.Children.Add(ctrlGrid);
        }
        private void saveOthCol_Click(object sender, RoutedEventArgs e)
        {
            List<string> newCols = new List<string>();

            foreach (var child in othContent.Children)
            {
                if (child is Grid grid)
                {
                    TextBox rTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("R"));
                    TextBox gTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("G"));
                    TextBox bTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("B"));

                    if (rTextBox != null && gTextBox != null && bTextBox != null)
                    {
                        if (int.TryParse(rTextBox.Text, out int r) && int.TryParse(gTextBox.Text, out int g) && int.TryParse(bTextBox.Text, out int b))
                        {
                            int rgb = (b << 16) | (g << 8) | r;
                            newCols.Add(rgb.ToString());

                            rTextBox.Text = "";
                            gTextBox.Text = "";
                            bTextBox.Text = "";
                        }
                    }
                }
            }
            saveFun(newCols);
        }
        private void resetOthCol_Click(object sender, RoutedEventArgs e)
        {
            othContent.Children.Clear();

            Label label = new Label
            {
                Content = selectedItem,
                Style = FindResource("ContentTitle") as Style,

            };
            othContent.Children.Add(label);

            //string jsonPath = System.IO.Path.Combine(baseDirectory, "defaultData.json");
            //string json = File.ReadAllText(jsonPath);
            //Dictionary<string, List<string>> defaultData = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);

            Dictionary<string, List<string>> defaultData = Globals.ThisAddIn.DefFunctionData;

            // clear and remake the content page
            string fun = (string)allFunctions[selectedItem]["fun"];

            // iterate over data stored in function and add each color to the appropriate textboxes
            int countForm = defaultData[fun].Count;

            for (int gridCount = 0; gridCount < countForm; gridCount++)
            {
                Grid newGrid = new Grid
                {
                    Name = "c" + (gridCount + 1).ToString(),
                };

                int rgb = int.Parse(defaultData[fun][gridCount]);

                int b = (rgb >> 16) & 255;
                int g = (rgb >> 8) & 255;
                int r = rgb & 255;

                Button colDis = new Button
                {
                    Style = FindResource("ColorDisplay") as Style,
                };
                colDis.Click += colButton_Click;

                TextBox newTB1 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "R", r.ToString(), new Thickness(0, 5, 110, 5));
                TextBox newTB2 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "G", g.ToString(), new Thickness(0, 5, 0, 5));
                TextBox newTB3 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "B", b.ToString(), new Thickness(110, 5, 0, 5));

                newGrid.Children.Add(colDis);
                newGrid.Children.Add(newTB1);
                newGrid.Children.Add(newTB2);
                newGrid.Children.Add(newTB3);

                othContent.Children.Add(newGrid);
                RGBTextBox(newGrid);
            }

            Grid ctrlGrid = new Grid
            {
                Margin = new Thickness(0, 20, 0, 0)
            };

            Button saveButton = new Button
            {
                Name = "saveCol",
                Content = "Save",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 0, 100, 0),
            };
            saveButton.Click += saveOthCol_Click;

            Button resetButton = new Button
            {
                Name = "resetCol",
                Content = "Reset",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(100, 0, 0, 0),
            };
            resetButton.Click += resetOthCol_Click;

            ctrlGrid.Children.Add(saveButton);
            ctrlGrid.Children.Add(resetButton);

            othContent.Children.Add(ctrlGrid);
        }
        //private void saveAutoCol_Click(object sender, RoutedEventArgs e)
        //{
        //    List<string> newCols = new List<string>();

        //    foreach (var child in othContent.Children)
        //    {
        //        if (child is Grid grid)
        //        {
        //            TextBox rTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("R"));
        //            TextBox gTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("G"));
        //            TextBox bTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("B"));

        //            if (rTextBox != null && gTextBox != null && bTextBox != null)
        //            {
        //                if (int.TryParse(rTextBox.Text, out int r) && int.TryParse(gTextBox.Text, out int g) && int.TryParse(bTextBox.Text, out int b))
        //                {
        //                    int rgb = (b << 16) | (g << 8) | r;
        //                    newCols.Add(rgb.ToString());

        //                    rTextBox.Text = "";
        //                    gTextBox.Text = "";
        //                    bTextBox.Text = "";
        //                }
        //            }
        //        }
        //    }
        //    saveFun(newCols);
        //}
        private void resetAutoCol_Click(object sender, RoutedEventArgs e)
        {
            othContent.Children.Clear();

            Label label = new Label
            {
                Content = selectedItem,
                Style = FindResource("ContentTitle") as Style,

            };
            othContent.Children.Add(label);

            //string jsonPath = System.IO.Path.Combine(baseDirectory, "defaultData.json");
            //string json = File.ReadAllText(jsonPath);
            //Dictionary<string, List<string>> defaultData = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);

            Dictionary<string, List<string>> defaultData = Globals.ThisAddIn.DefFunctionData;

            // clear and remake the content page
            string fun = (string)allFunctions[selectedItem]["fun"];

            // iterate over data stored in function and add each color to the appropriate textboxes
            int countForm = defaultData[fun].Count;
            List<string> labels = new List<string> { "Hardcode", "Partial Input", "Formula", "Sheet Reference", "File Reference" };

            for (int gridCount = 0; gridCount < countForm; gridCount++)
            {
                Grid newGrid = new Grid
                {
                    Name = "c" + (gridCount + 1).ToString(),
                };

                int rgb = int.Parse(defaultData[fun][gridCount]);

                int b = (rgb >> 16) & 255;
                int g = (rgb >> 8) & 255;
                int r = rgb & 255;

                Button colDis = new Button
                {
                    Style = FindResource("ColorDisplay") as Style,
                };
                colDis.Click += colButton_Click;

                TextBox newTB1 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "R", r.ToString(), new Thickness(0, 5, 110, 5));
                TextBox newTB2 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "G", g.ToString(), new Thickness(0, 5, 0, 5));
                TextBox newTB3 = CreateColorTextBox("c" + (gridCount + 1).ToString() + "B", b.ToString(), new Thickness(110, 5, 0, 5));

                TextBlock colLabel = new TextBlock
                {
                    Style = FindResource("FunctionName") as Style,
                    Margin = new Thickness(290, 5, 0, 5),
                    Text = labels[gridCount],
                    HorizontalAlignment = HorizontalAlignment.Left
                };

                newGrid.Children.Add(colDis);
                newGrid.Children.Add(newTB1);
                newGrid.Children.Add(newTB2);
                newGrid.Children.Add(newTB3);
                newGrid.Children.Add(colLabel);

                othContent.Children.Add(newGrid);
                RGBTextBox(newGrid);
            }

            Grid ctrlGrid = new Grid
            {
                Margin = new Thickness(0, 20, 0, 0)
            };

            Button saveButton = new Button
            {
                Name = "saveCol",
                Content = "Save",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(0, 0, 100, 0),
            };
            saveButton.Click += saveOthCol_Click;

            Button resetButton = new Button
            {
                Name = "resetCol",
                Content = "Reset",
                Style = FindResource("AddButtonStyle") as Style,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 80,
                Margin = new Thickness(100, 0, 0, 0),
            };
            resetButton.Click += resetOthCol_Click;

            ctrlGrid.Children.Add(saveButton);
            ctrlGrid.Children.Add(resetButton);

            othContent.Children.Add(ctrlGrid);
        }
        public void saveFun(List<string> data)
        {
            string fun = (string)allFunctions[selectedItem]["fun"];

            userData[fun] = data; // new, processed info into list form
            string functionsJson = JsonConvert.SerializeObject(userData, Formatting.Indented);
            File.WriteAllText(Globals.ThisAddIn.FunctionDataPath, functionsJson);

            if (data.Count == 0)
            {
                shortcutsData[fun]["enabled"][0] = "false";

                try
                {
                    string shortcutsJson = JsonConvert.SerializeObject(shortcutsData, Formatting.Indented);
                    File.WriteAllText(System.IO.Path.Combine(Globals.ThisAddIn.appData, "shortcuts.json"), shortcutsJson);
                }
                catch (Exception ex)
                {
                    Globals.ThisAddIn.ShowErrorMessage(ex.Message);
                }
            }

            string json = File.ReadAllText(Globals.ThisAddIn.FunctionDataPath);
            userData = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);

            this.Close();
            callerForm.FunctionList.SelectedItem = null;
        }
        public void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            Button clickedButton = (Button)sender;
            Grid parentGrid = (Grid)clickedButton.Parent;
            if (numFunctions.Contains(selectedItem))
            {
                numContent.Children.Remove(parentGrid);
            }
            else if (colFunctions.Contains(selectedItem))
            {
                colContent.Children.Remove(parentGrid);
            };
        }
        private TextBox CreateColorTextBox(string name, string text, Thickness thickness)
        {
            TextBox newTB = new TextBox
            {
                Name = name,
                Text = text,
                Style = FindResource("ColorTextBox") as Style,
                Margin = thickness
            };
            newTB.TextChanged += OnRGBTextChanged;
            newTB.PreviewTextInput += ColTextLimit;

            newTB.ApplyTemplate();

            // Assuming the RepeatButtons were named PART_IncreaseButton and PART_DecreaseButton in your style
            RepeatButton increaseButton = (RepeatButton)newTB.Template.FindName("PART_IncreaseButton", newTB);
            RepeatButton decreaseButton = (RepeatButton)newTB.Template.FindName("PART_DecreaseButton", newTB);

            increaseButton.Click += IncreaseButton_Click;
            decreaseButton.Click += DecreaseButton_Click;

            return newTB;
        }
        public void OnRGBTextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tb = (TextBox)sender;
            Grid grid = (Grid)tb.Parent;

            TextBox rTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("R"));
            TextBox gTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("G"));
            TextBox bTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("B"));

            // Parse RGB values from TextBoxes
            if (!byte.TryParse(rTextBox.Text, out byte red)) red = 0;
            if (!byte.TryParse(gTextBox.Text, out byte green)) green = 0;
            if (!byte.TryParse(bTextBox.Text, out byte blue)) blue = 0;

            // Update Rectangle fill
            //grid.Children.OfType<System.Windows.Shapes.Rectangle>().FirstOrDefault().Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(red, green, blue));
            grid.Children.OfType<Button>().FirstOrDefault().Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(red, green, blue));
        }
        public Color RGBTextBox(Grid grid)
        {
            TextBox rTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("R"));
            TextBox gTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("G"));
            TextBox bTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("B"));

            // Parse RGB values from TextBoxes
            if (!byte.TryParse(rTextBox.Text, out byte red)) red = 0;
            if (!byte.TryParse(gTextBox.Text, out byte green)) green = 0;
            if (!byte.TryParse(bTextBox.Text, out byte blue)) blue = 0;

            // Update Rectangle fill
            //grid.Children.OfType<System.Windows.Shapes.Rectangle>().FirstOrDefault().Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(red, green, blue));
            grid.Children.OfType<Button>().FirstOrDefault().Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(red, green, blue));
            return Color.FromRgb(red, green, blue);
        }
        public void ColTextLimit(object sender, TextCompositionEventArgs e)
        {
            if (!int.TryParse(e.Text, out int value))
            {
                e.Handled = true; // Discard non-numeric input
            }
            else
            {
                TextBox tb = (TextBox)sender;
                // Adjust the minimum and maximum values as per your requirements
                int minValue = 0;
                int maxValue = 255;

                // Get the current value of the TextBox
                if (int.TryParse(tb.Text + e.Text, out int newValue))
                {
                    // Check if the new value is within the specified range
                    if (newValue < minValue || newValue > maxValue)
                    {
                        e.Handled = true; // Discard the input if it's outside the range
                    }
                }
            }
        }
        public void IncreaseValue(TextBox textBox)
        {
            if (textBox != null)
            {
                if (int.TryParse(textBox.Text, out int currentValue))
                {
                    if (currentValue < 255)
                    {
                        textBox.Text = (currentValue + 1).ToString();
                    }
                }
                else
                {
                    textBox.Text = 1.ToString();
                }
            }
        }
        public void DecreaseValue(TextBox textBox)
        {
            if (textBox != null)
            {
                if (int.TryParse(textBox.Text, out int currentValue))
                {
                    if (currentValue > 0)
                    {
                        textBox.Text = (currentValue - 1).ToString();

                    }
                }
                else
                {
                    textBox.Text = 1.ToString();
                }
            }
        }
        private void IncreaseButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement frameworkElement && frameworkElement.TemplatedParent is TextBox textBox)
            {
                IncreaseValue(textBox);
            }
        }
        private void DecreaseButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement frameworkElement && frameworkElement.TemplatedParent is TextBox textBox)
            {
                DecreaseValue(textBox);
            }
        }
        public void colButton_Click(object sender, RoutedEventArgs e)
        {
            // Retrieve the button that was clicked
            Button clickedButton = sender as Button;

            Grid parent = (Grid)clickedButton.Parent;
            Color color = RGBTextBox(parent);

            // Show the color selection dialog
            colorFormPopup = ShowColorWindow(color);

            if (colorFormPopup.DialogResult == true)
            {
                // Extract RGB values from the selected color
                Color selectedColor = colorFormPopup.SelectedColor;

                // Set the RGB values
                string r = selectedColor.R.ToString();
                string g = selectedColor.G.ToString();
                string b = selectedColor.B.ToString();

                Button button = (Button)sender;
                Grid grid = (Grid)button.Parent;

                // Get textboxes by their names
                TextBox rTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("R"));
                TextBox gTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("G"));
                TextBox bTextBox = grid.Children.OfType<TextBox>().FirstOrDefault(t => t.Name.EndsWith("B"));

                // Set the RGB values
                rTextBox.Text = r;
                gTextBox.Text = g;
                bTextBox.Text = b;

                // Parse RGB values from TextBoxes
                if (!byte.TryParse(r, out byte red)) red = 0;
                if (!byte.TryParse(g, out byte green)) green = 0;
                if (!byte.TryParse(b, out byte blue)) blue = 0;

                // Update Rectangle fill
                //grid.Children.OfType<System.Windows.Shapes.Rectangle>().FirstOrDefault().Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(red, green, blue));
                grid.Children.OfType<Button>().FirstOrDefault().Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(red, green, blue));
            }
        }
        public ColorFormPopup ShowColorWindow(Color color)
        {
            //if (colorFormPopup != null && !colorFormPopup.IsClosed)
            //{
            //    colorFormPopup.Close();
            //    colorFormPopup = null;
            //}

            colorFormPopup = new ColorFormPopup();

            // Set the owner of the WPF window to the Excel application.
            WindowInteropHelper windowInteropHelper = new WindowInteropHelper(colorFormPopup);
            windowInteropHelper.Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd);

            colorFormPopup.SelectedColor = color;

            colorFormPopup.ShowDialog();
            //SetForegroundWindow(windowInteropHelper.Handle);

            return colorFormPopup;
        }
    }
}
