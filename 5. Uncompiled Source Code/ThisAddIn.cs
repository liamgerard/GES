using GlobalHotKey;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace GES
{
    public partial class ThisAddIn
    {
        #region Startup/Shutdown
        public HotKeyManager hotKeyManager;
        public Dictionary<HotKey, System.Action> hotKeyActions = new Dictionary<HotKey, System.Action>();
        public bool verifyCheck = false;

        #region Public Data
        public Dictionary<string, List<string>> FunctionData;
        public Dictionary<string, Dictionary<string, List<string>>> ShortcutsData;
        public Dictionary<string, List<string>> DefFunctionData;
        public Dictionary<string, Dictionary<string, List<string>>> DefShortcutsData;
        public Dictionary<string, Dictionary<string, object>> Functions;
        public Dictionary<string, string> KeysToEnum;
        public Dictionary<string, string> EnumToKeys;
        public Dictionary<string, List<string>> DisabledKeys;

        public string ShortcutsPath;
        public string FunctionDataPath;
        public string DisabledKeysPath;

        public string appData;

        public int maxCellsFast = 20001;
        public int maxCellsSlow = 5001;
        public int maxCells = 50000;
        #endregion

        public Microsoft.Office.Tools.CustomTaskPane shortcutsTaskPane;
        public Microsoft.Office.Tools.CustomTaskPane formatsTaskPane;
        public Microsoft.Office.Tools.CustomTaskPane disabledKeysTaskpane;

        [Obfuscation(Exclude = true)]
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(true);
            Settings_Startup();
        }

        [Obfuscation(Exclude = true)]
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Settings_Shutdown();
        }
        public void Settings_Startup()
        {
            #region File Data Region
            string genAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            appData = Path.Combine(genAppDataFolder, "GES");
            if (!Directory.Exists(appData))
            {
                Directory.CreateDirectory(appData);
            }

            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            FunctionDataPath = Path.Combine(appData, "userData.json");
            if (!File.Exists(FunctionDataPath)) File.Copy(Path.Combine(baseDirectory, "userData.json"), FunctionDataPath);
            this.FunctionData = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(File.ReadAllText(FunctionDataPath));

            ShortcutsPath = Path.Combine(appData, "shortcuts.json");
            if (!File.Exists(ShortcutsPath)) File.Copy(Path.Combine(baseDirectory, "shortcuts.json"), ShortcutsPath);
            this.ShortcutsData = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, List<string>>>>(File.ReadAllText(ShortcutsPath));

            DisabledKeysPath = Path.Combine(appData, "disabledKeys.json");
            if (!File.Exists(DisabledKeysPath)) File.Copy(Path.Combine(baseDirectory, "disabledKeys.json"), DisabledKeysPath);
            this.DisabledKeys = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(File.ReadAllText(DisabledKeysPath));

            this.DefFunctionData = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(File.ReadAllText(Path.Combine(baseDirectory, "defaultData.json")));
            this.DefShortcutsData = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, List<string>>>>(File.ReadAllText(Path.Combine(baseDirectory, "Defaultshortcuts.json")));
            this.Functions = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, object>>>(File.ReadAllText(Path.Combine(baseDirectory, "functions.json")));

            this.KeysToEnum = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(Path.Combine(baseDirectory, "keysToEnum.json")));
            this.EnumToKeys = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(Path.Combine(baseDirectory, "enumToKeys.json")));
            #endregion

            #region Check For New Data
            List<string> keysToAdd = new List<string>();

            foreach (string key in DefFunctionData.Keys)
            {
                if (!FunctionData.Keys.Contains(key))
                {
                    keysToAdd.Add(key);
                }
            }

            foreach (string key in keysToAdd)
            {
                FunctionData.Add(key, DefFunctionData[key]);
            }

            keysToAdd = new List<string>();

            foreach (string key in DefShortcutsData.Keys)
            {
                if (!ShortcutsData.Keys.Contains(key))
                {
                    keysToAdd.Add(key);
                }
            }

            foreach (string key in keysToAdd)
            {
                ShortcutsData.Add(key, DefShortcutsData[key]);
            }

            File.WriteAllText(Globals.ThisAddIn.FunctionDataPath, JsonConvert.SerializeObject(FunctionData, Formatting.Indented));
            File.WriteAllText(Globals.ThisAddIn.ShortcutsPath, JsonConvert.SerializeObject(ShortcutsData, Formatting.Indented));

            #endregion

            System.Windows.Forms.UserControl shortcutsMenu = new UserControl1();
            shortcutsTaskPane = this.CustomTaskPanes.Add(shortcutsMenu, "Shortcuts Menu");
            shortcutsTaskPane.Width = 675;
            shortcutsTaskPane.Visible = false;

            System.Windows.Forms.UserControl formatMenu = new UserControl2();
            formatsTaskPane = this.CustomTaskPanes.Add(formatMenu, "Formatting Menu");
            formatsTaskPane.Width = 300;
            formatsTaskPane.Visible = false;

            System.Windows.Forms.UserControl disabledKeysMenu = new UserControl3();
            disabledKeysTaskpane = this.CustomTaskPanes.Add(disabledKeysMenu, "Disabled Keys Menu");
            disabledKeysTaskpane.Width = 300;
            disabledKeysTaskpane.Visible = false;

            hotKeyManager = new HotKeyManager();

            #region Hotkey Registration Region

            foreach (var shortcutData in this.ShortcutsData)
            {
                string functionName = shortcutData.Key;
                var shortcutInfo = shortcutData.Value;

                if (bool.Parse(shortcutInfo["enabled"][0]))
                {
                    // get main key
                    string mainKey = shortcutInfo["mainKey"][0];
                    Key convertedMainKey = (Key)Enum.Parse(typeof(Key), mainKey);

                    // get modifier keys
                    List<string> modKeys = shortcutInfo["modKeys"];
                    ModifierKeys compoundModifiers = ModifierKeys.None;

                    // compound the modifier keys
                    foreach (string modifierKey in modKeys)
                    {
                        compoundModifiers |= (ModifierKeys)Enum.Parse(typeof(ModifierKeys), modifierKey);
                    }

                    RegisterHotKeyAction(convertedMainKey, compoundModifiers, GetActionForFunctionName(functionName));
                }
            }
            foreach (var disKey in this.DisabledKeys["keys"])
            {
                Key convertedDisKey = (Key)Enum.Parse(typeof(Key), disKey);
                RegisterDisabledKeyAction(convertedDisKey, BlankFunction);
            }
            #endregion

            hotKeyManager.KeyPressed += HotKeyManagerPressed;
        }
        public void Settings_Shutdown()
        {
            for (int i = Globals.ThisAddIn.CustomTaskPanes.Count - 1; i >= 0; i--)
            {
                var taskpane = Globals.ThisAddIn.CustomTaskPanes[i];
                Globals.ThisAddIn.CustomTaskPanes.Remove(taskpane);
            }

            // Unregister all hotkeys.
            foreach (var hotKey in hotKeyActions.Keys)
            {
                hotKeyManager.Unregister(hotKey);
            }
        }
        public void RegisterHotKeyAction(Key key, ModifierKeys modifiers, System.Action action)
        {
            //var hotKey = hotKeyManager.Register(key, modifiers);
            //hotKeyActions[hotKey] = action;
            try
            {
                var hotKey = hotKeyManager.Register(key, modifiers);
                hotKeyActions[hotKey] = action;
            }
            catch //(System.ComponentModel.Win32Exception ex)
            {
            }
        }
        public void RegisterDisabledKeyAction(Key key, System.Action action)
        {
            //var hotKey = hotKeyManager.Register(key, modifiers);
            //hotKeyActions[hotKey] = action;
            try
            {
                var hotKey = hotKeyManager.Register(key, ModifierKeys.None);
                hotKeyActions[hotKey] = action;
            }
            catch //(System.ComponentModel.Win32Exception ex)
            {
            }
        }
        public void HotKeyManagerPressed(object sender, KeyPressedEventArgs e)
        {
            // Look up the action for the hotkey that was pressed, and execute it.
            if (IsExcelActive() && hotKeyActions.TryGetValue(e.HotKey, out var action))
            {
                action();
            }
        }
        public Action GetActionForFunctionName(string functionName)
        {
            // Assuming the functions are defined in the same class, you can use reflection to obtain the MethodInfo for the function
            MethodInfo methodInfo = GetType().GetMethod(functionName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);
            if (methodInfo != null)
            {
                // Create a delegate for the function using the MethodInfo
                return (Action)Delegate.CreateDelegate(typeof(Action), this, methodInfo);
            }
            else
            {
                // Handle the case where the function is not found
                Console.WriteLine($"Function '{functionName}' not found");
                return null;
            }
        }
        #endregion

        #region Obfuscated Functions
        private void ToggleCellSign__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (TooManyCells(maxCellsSlow)) return;
                int count = 0;

                foreach (Excel.Range cell in selection.Cells)
                {
                    //stop after X cells
                    count += 1;
                    if (count >= maxCellsFast) return;

                    if (cell.Value == null)
                    {
                        continue;
                    }

                    // just value
                    if (!cell.HasFormula && decimal.TryParse(cell.Value2.ToString(), out decimal value))
                    {
                        cell.Value = cell.Value * -1;
                        continue;
                    }

                    // formula
                    string form = cell.Formula;
                    if (string.IsNullOrWhiteSpace(form) || form[0] != '=') // Check if formula starts with '='
                    {
                        continue;
                    }

                    // has non-negative formula
                    if (!(form.StartsWith("=-(") && form.EndsWith(")")))
                    {
                        form = "=-(" + form.Substring(1) + ")";
                        cell.Formula = form;
                        continue;
                    }

                    // check parentheses
                    // copy formula, check if formula would cause error due to unmatched parentheses
                    Stack<char> stack = new Stack<char>();
                    bool unmatched = false;
                    string form1 = "=" + form.Substring(3, form.Length - 4);

                    for (int i = 0; i < form1.Length; i++)
                    {
                        if (form1[i] == '(')
                        {
                            stack.Push('(');
                        }
                        else if (form1[i] == ')')
                        {
                            if (stack.Count == 0)
                            {
                                unmatched = true;
                                break;
                            }
                            else
                            {
                                stack.Pop();
                            }
                        }
                    }

                    // check for unmatched opening parentheses
                    // if unmatched == true or there are remaining parentheses in the stack
                    unmatched = unmatched || stack.Count > 0;

                    // if formula would be unmatched, make negative. Otherwise, make positive
                    if (unmatched)
                    {
                        form = "=-(" + form.Substring(1) + ")";
                    }
                    else
                    {
                        form = "=" + form.Substring(3, form.Length - 4);
                    }
                    cell.Formula = form;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ToggleFontColor__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Font == null || selection == null) return;

                int color = System.Convert.ToInt32(activeCell.Font.Color);

                List<int> data = this.FunctionData["ToggleFontColor"].Select(str => int.Parse(str)).ToList();
                int currentIndex = data.IndexOf(color);

                color = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.Font.Color = color;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleFontColor1__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Font == null || selection == null) return;

                int color = System.Convert.ToInt32(activeCell.Font.Color);

                List<int> data = this.FunctionData["CycleFontColor1"].Select(str => int.Parse(str)).ToList();
                int currentIndex = data.IndexOf(color);

                color = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.Font.Color = color;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleFontColor2__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Font == null || selection == null) return;

                int color = System.Convert.ToInt32(activeCell.Font.Color);

                List<int> data = this.FunctionData["CycleFontColor2"].Select(str => int.Parse(str)).ToList();
                int currentIndex = data.IndexOf(color);

                color = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.Font.Color = color;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void AutoColorCells__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                List<int> data = this.FunctionData["AutoColorCells"].Select(str => int.Parse(str)).ToList();

                if (TooManyCells(maxCellsFast)) { return; }
                int count = 0;

                foreach (Excel.Range cell in selection.Cells)
                {
                    // stop after x cells
                    count += 1;
                    if (count >= maxCellsFast) return;

                    if (cell == null || cell.Font == null || string.IsNullOrWhiteSpace(cell.Formula)) continue;

                    string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    char[] seps = new char[] { '+', '-', '*', '/', '^', '&', '(', '=', ',', ':', ')', '>', '<' };
                    string formula = cell.Formula;
                    List<double> formList = new List<double> { }; // double is okay because Excel uses double in cells
                    bool cellRef = false;
                    bool partialInput = false;

                    if (cell.HasFormula)
                    {
                        cellRef = false;

                        if (formula == null)
                        {
                            continue;
                        }

                        for (int i = 0; i < formula.Length; i++)
                        {
                            if (letters.Contains(formula.Substring(i, 1)))
                            {
                                cellRef = true;
                            }
                        }

                        formList = formula.Replace(" ", "")
                                            .Split(seps)
                                            .Where(str => double.TryParse(str, out _))
                                            .Select(str => double.Parse(str))
                                            .ToList();

                        for (int i = 0; i < formList.Count;)
                        {
                            double logResult = Math.Log10(Math.Abs(formList[i]));
                            if (Math.Abs(logResult - Math.Round(logResult)) < 1E-10) formList.RemoveAt(i);
                            else i++;
                        }

                        if (formList.Count > 0)
                        {
                            partialInput = true;
                        }

                        if (formula.Contains("!"))
                        {
                            if (formula.Contains("["))
                            {
                                cell.Font.Color = data[4]; // link to file
                            }
                            else
                            {
                                cell.Font.Color = data[3]; // link to sheet
                            }
                        }
                        else if (partialInput)
                        {
                            if (cellRef)
                            {
                                cell.Font.Color = data[1]; // partial input
                            }
                            else
                            {
                                cell.Font.Color = data[0]; // math formula
                            }
                        }
                        else if (cellRef)
                        {
                            cell.Font.Color = data[2]; // cell reference formula
                        }
                    }
                    else if (decimal.TryParse(formula, out decimal value))
                    {
                        cell.Font.Color = data[0]; // hardcoded cell - no "="
                    }
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleFillColor2__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Font == null || activeCell?.Interior == null || selection == null || selection?.Interior == null) return;

                int color = System.Convert.ToInt32(activeCell.Interior.Color);
                bool noFill = (Excel.XlPattern)activeCell.Interior.Pattern == Excel.XlPattern.xlPatternNone;

                List<int> data = this.FunctionData["CycleFillColor2"].Select(str => int.Parse(str)).ToList();
                int currentIndex = data.IndexOf(color);

                if (noFill && currentIndex == 0)
                {
                    selection.Interior.Color = data[(currentIndex) % data.Count]; // mod makes sure it wraps back to zero
                }
                else if (noFill)
                {
                    selection.Interior.Color = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero
                }
                else if (currentIndex == (data.Count - 1))
                {
                    selection.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                }
                else
                {
                    selection.Interior.Color = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleFillColor1__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Font == null || activeCell?.Interior == null || selection == null || selection?.Interior == null) return;

                if (activeCell != null)
                {
                    int color = System.Convert.ToInt32(activeCell.Interior.Color);
                    bool noFill = (Excel.XlPattern)activeCell.Interior.Pattern == Excel.XlPattern.xlPatternNone;

                    List<int> data = this.FunctionData["CycleFillColor1"].Select(str => int.Parse(str)).ToList();
                    int currentIndex = data.IndexOf(color);

                    if (noFill && currentIndex == 0)
                    {
                        selection.Interior.Color = data[(currentIndex) % data.Count]; // mod makes sure it wraps back to zero
                    }
                    else if (noFill)
                    {
                        selection.Interior.Color = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero
                    }
                    else if (currentIndex == (data.Count - 1))
                    {
                        selection.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                    }
                    else
                    {
                        selection.Interior.Color = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero
                    }
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void IncreaseDecimal__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCellsFast)) { return; }

                string param = activeCell.NumberFormat;

                if (param == "General")
                {
                    param = "0";
                }

                int countDec = param.Length - param.Replace(".", "").Length;
                int countSemi = param.Length - param.Replace(";", "").Length;

                // Split format into sections
                string[] sections = param.Split(';');

                for (int i = 0; i < sections.Length; i++)
                {
                    // Check if section contains a decimal
                    if (!sections[i].Contains("."))
                    {
                        // If not, add decimal and zero to end of numeric part
                        int index = sections[i].LastIndexOfAny("0123456789".ToCharArray());
                        if (index != -1)
                        {
                            sections[i] = sections[i].Insert(index + 1, ".0");
                        }
                    }
                    else
                    {
                        // If does, add zero to end of decimal part
                        int index = sections[i].IndexOf('.');
                        if (index != -1)
                        {
                            sections[i] = sections[i].Insert(index + 2, "0");
                        }
                    }
                }

                // Join sections
                param = string.Join(";", sections);

                selection.NumberFormat = param;

                //foreach (Excel.Range cell in selection)
                //{
                //    cell.NumberFormat = param;
                //}
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void DecreaseDecimal__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCellsFast)) { return; }

                string param = activeCell.NumberFormat;

                // Split format into sections
                string[] sections = param.Split(';');

                for (int i = 0; i < sections.Length; i++)
                {
                    // Check if section contains a decimal
                    if (sections[i].Contains("."))
                    {
                        // Find position of decimal
                        int decimalIndex = sections[i].IndexOf('.');

                        // Check if zero immediately after the decimal
                        if (decimalIndex + 1 < sections[i].Length && sections[i][decimalIndex + 1] == '0')
                        {
                            // Remove the zero
                            sections[i] = sections[i].Remove(decimalIndex + 1, 1);
                        }

                        // If no more digits after decimal, remove it
                        if (decimalIndex + 1 >= sections[i].Length || !char.IsDigit(sections[i][decimalIndex + 1]))
                        {
                            sections[i] = sections[i].Remove(decimalIndex, 1);
                        }
                    }
                }

                // Join sections
                param = string.Join(";", sections);

                selection.NumberFormat = param;

                //foreach (Excel.Range cell in selection)
                //{
                //    cell.NumberFormat = param;
                //}
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ShiftDecimalLeft__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (TooManyCells(maxCellsSlow)) { return; }
                int count = 0;

                foreach (Excel.Range cell in selection.Cells)
                {
                    // stop after X cells
                    count += 1;
                    if (count >= maxCellsSlow) return;

                    if (cell.Value == null)
                    {
                        continue;
                    }

                    if (cell.HasFormula)
                    {
                        string form = cell.Formula;
                        string formCopy = cell.Formula;
                        decimal valCopy = (decimal)cell.Value / 10;

                        if (form.StartsWith("=(10"))
                        {
                            form = "=(1" + form.Substring(4);
                        }
                        else if (form.StartsWith("=(0."))
                        {
                            form = "=(0.0" + form.Substring(4);
                        }
                        else
                        {
                            if (!form.StartsWith("=(0.1)*("))
                            {
                                form = "=(0.1)*(" + form.Substring(1) + ")";
                            }
                        }
                        if (form.StartsWith("=(1)*("))
                        {
                            //form = "=" + form.Substring(6, form.Length - 7);

                            Stack<char> stack = new Stack<char>();
                            bool unmatched = false;
                            string form1 = "=" + form.Substring(6, form.Length - 7);

                            for (int i = 0; i < form1.Length; i++)
                            {
                                if (form1[i] == '(')
                                {
                                    stack.Push('(');
                                }
                                else if (form1[i] == ')')
                                {
                                    if (stack.Count == 0)
                                    {
                                        unmatched = true;
                                        break;
                                    }
                                    else
                                    {
                                        stack.Pop();
                                    }
                                }
                            }

                            // check for unmatched opening parentheses
                            // if unmatched == true or there are remaining parentheses in the stack
                            unmatched = unmatched || stack.Count > 0;

                            if (unmatched)
                            {
                                // if removing the "=(1)*(" would cause unmatched, dont change it
                                form = "=(0.1)*((10)*(" + form.Substring(6, form.Length - 6) + ")";
                            }
                            else
                            {
                                form = "=" + form.Substring(6, form.Length - 7);
                            }
                        }
                        cell.Formula = form;
                        decimal val12121 = Math.Abs((decimal)cell.Value);
                        decimal val122113 = (decimal)valCopy;
                        if (Math.Abs((decimal)cell.Value - (decimal)valCopy) > (decimal)1e-15)
                        {
                            cell.Formula = "=(0.1)*(" + formCopy.Substring(1) + ")";
                        }
                    }
                    else if (decimal.TryParse(cell.Value2.ToString(), out decimal value))
                    {
                        cell.Value2 = value / 10;
                    }
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ShiftDecimalRight__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (TooManyCells(maxCellsSlow)) return;
                int count = 0;

                foreach (Excel.Range cell in selection.Cells)
                {
                    // stop after X cells
                    count += 1;
                    if (count >= maxCellsSlow) return;

                    if (cell.Value == null)
                    {
                        continue;
                    }

                    if (cell.HasFormula)
                    {
                        string form = cell.Formula;
                        string formCopy = cell.Formula;
                        decimal valCopy = (decimal)cell.Value * 10;
                        if (form.StartsWith("=(0.1)*("))
                        {
                            //form = "=" + form.Substring(8, form.Length - 9);

                            Stack<char> stack = new Stack<char>();
                            bool unmatched = false;
                            string form1 = "=" + form.Substring(8, form.Length - 9);

                            for (int i = 0; i < form1.Length; i++)
                            {
                                if (form1[i] == '(')
                                {
                                    stack.Push('(');
                                }
                                else if (form1[i] == ')')
                                {
                                    if (stack.Count == 0)
                                    {
                                        unmatched = true;
                                        break;
                                    }
                                    else
                                    {
                                        stack.Pop();
                                    }
                                }
                            }

                            // check for unmatched opening parentheses
                            // if unmatched == true or there are remaining parentheses in the stack
                            unmatched = unmatched || stack.Count > 0;

                            if (unmatched)
                            {
                                form = "=(10)*(" + form.Substring(1, form.Length - 1) + ")";
                            }
                            else
                            {
                                form = "=" + form.Substring(8, form.Length - 9);
                            }
                        }
                        else if (form.StartsWith("=(0.0"))
                        {
                            form = "=(0." + form.Substring(5, form.Length - 5);
                        }
                        else if (form.StartsWith("=(1")) // && form.Contains(".")) 
                        {
                            form = "=(10" + form.Substring(3, form.Length - 3);
                        }
                        else
                        {
                            form = "=(10)*(" + form.Substring(1, form.Length - 1) + ")";
                        }
                        cell.Formula = form;
                        decimal val12121 = Math.Abs((decimal)cell.Value);
                        decimal val122113 = (decimal)valCopy;
                        if (Math.Abs((decimal)cell.Value - (decimal)valCopy) > (decimal)1e-15)
                        {
                            cell.Formula = "=(10)*(" + formCopy.Substring(1) + ")";
                        }
                    }
                    else if (decimal.TryParse(cell.Value2.ToString(), out decimal value))
                    {
                        cell.Value2 = value * 10;
                    }
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ToggleGeneralNumberFormats__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["ToggleGeneralNumberFormats"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleDateFormats__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["CycleDateFormats"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleCurrency__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["CycleCurrency"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleForeignCurrency__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                // dollar $
                // euro €
                // yen ¥
                // pound £
                // rupee ₹
                // bitcoin ₿

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["CycleForeignCurrency"];

                foreach (string curr in data)
                {
                    if (param.Contains(curr)) // if param contains the currency
                    {
                        int currentIndex = data.IndexOf(curr);
                        param = param.Replace(curr, data[(currentIndex + 1) % data.Count]); // then replace the currency with the next currency in the list
                        break;
                    }
                }

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CyclePercent__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["CyclePercent"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ToggleBPS__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                if (TooManyCells(maxCellsSlow)) return;

                string param = activeCell.NumberFormat;
                if (param != "#,##0\"bps\"_);(#,##0\"bps\");@")
                {
                    param = "#,##0\"bps\"_);(#,##0\"bps\");@";
                    int count = 0;
                    foreach (Excel.Range cell in selection.Cells)
                    {
                        // stop after X cells
                        count += 1;
                        if (count >= maxCellsSlow)
                        {
                            return;
                        }

                        if (cell.Value == null)
                        {
                            continue;
                        }

                        // just value
                        if (!cell.HasFormula && decimal.TryParse(cell.Value2.ToString(), out decimal value))
                        {
                            cell.Value = cell.Value * 10000;
                            continue;
                        }

                        // formula
                        string form = cell.Formula;
                        if (string.IsNullOrWhiteSpace(form) || form[0] != '=') // Check if formula starts with '='
                        {
                            continue;
                        }

                        // check parentheses
                        // copy formula, check if formula would cause error due to unmatched parentheses
                        Stack<char> stack = new Stack<char>();
                        bool unmatched = false;

                        if (form.StartsWith("=(0.0001)*("))
                        {
                            string form1 = "=" + form.Substring(11, form.Length - 12);

                            for (int i = 0; i < form1.Length; i++)
                            {
                                if (form1[i] == '(')
                                {
                                    stack.Push('(');
                                }
                                else if (form1[i] == ')')
                                {
                                    if (stack.Count == 0)
                                    {
                                        unmatched = true;
                                        break;
                                    }
                                    else
                                    {
                                        stack.Pop();
                                    }
                                }
                            }
                        }
                        // check for unmatched opening parentheses
                        // if unmatched == true or there are remaining parentheses in the stack
                        unmatched = unmatched || stack.Count > 0;

                        // if formula would be unmatched, add .0001. Otherwise, remove "=(10000)*("
                        if (unmatched || !form.StartsWith("=(0.0001)*("))
                        {
                            form = "=(10000)*(" + form.Substring(1) + ")";
                        }
                        else
                        {
                            form = "=" + form.Substring(11, form.Length - 12);
                        }
                        cell.Formula = form;
                    }
                }
                else
                {
                    param = "#,##0.0_);(#,##0.0)";
                    int count = 0;
                    foreach (Excel.Range cell in selection.Cells)
                    {
                        // stop after X cells
                        count += 1;
                        if (count >= maxCellsSlow)
                        {
                            return;
                        }

                        if (cell.Value == null)
                        {
                            continue;
                        }

                        // just value
                        if (!cell.HasFormula && decimal.TryParse(cell.Value2.ToString(), out decimal value))
                        {
                            cell.Value = cell.Value / 10000;
                            continue;
                        }

                        // formula
                        string form = cell.Formula;
                        if (string.IsNullOrWhiteSpace(form) || form[0] != '=') // Check if formula starts with '='
                        {
                            continue;
                        }

                        // check parentheses
                        // copy formula, check if formula would cause error due to unmatched parentheses
                        Stack<char> stack = new Stack<char>();
                        bool unmatched = false;

                        if (form.StartsWith("=(10000)*("))
                        {
                            string form1 = "=" + form.Substring(10, form.Length - 11);

                            for (int i = 0; i < form1.Length; i++)
                            {
                                if (form1[i] == '(')
                                {
                                    stack.Push('(');
                                }
                                else if (form1[i] == ')')
                                {
                                    if (stack.Count == 0)
                                    {
                                        unmatched = true;
                                        break;
                                    }
                                    else
                                    {
                                        stack.Pop();
                                    }
                                }
                            }
                        }

                        // check for unmatched opening parentheses
                        // if unmatched == true or there are remaining parentheses in the stack
                        unmatched = unmatched || stack.Count > 0;

                        // if formula would be unmatched, add .0001. Otherwise, remove "=(10000)*("
                        if (unmatched || !form.StartsWith("=(10000)*("))
                        {
                            form = "=(0.0001)*(" + form.Substring(1) + ")";
                        }
                        else
                        {
                            form = "=" + form.Substring(10, form.Length - 11);
                        }
                        cell.Formula = form;
                    }
                }
                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ToggleMultiple__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["ToggleMultiple"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ToggleBinary__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                // use 1 and 0
                // Y
                // Yes
                // On
                // True

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["ToggleBinary"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleOtherNumberFormats1__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["CycleOtherNumberFormats1"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleOtherNumberFormats2__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["CycleOtherNumberFormats2"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CycleOtherNumberFormats3__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.NumberFormat == null || selection == null) return;
                //if (TooManyCells(maxCells)) return;

                string param = activeCell.NumberFormat;

                List<string> data = this.FunctionData["CycleOtherNumberFormats3"];
                int currentIndex = data.IndexOf(param);

                param = data[(currentIndex + 1) % data.Count]; // mod makes sure it wraps back to zero

                selection.NumberFormat = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void UpperBorderRange__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null || selection?.Borders == null) return;

                Excel.Borders borders = selection.Borders;
                Excel.Border param = borders[Excel.XlBordersIndex.xlEdgeTop];

                if ((Excel.XlLineStyle)param.LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                {
                    param.LineStyle = Excel.XlLineStyle.xlDot;
                }
                else if ((Excel.XlLineStyle)param.LineStyle == Excel.XlLineStyle.xlDot)
                {
                    param.LineStyle = Excel.XlLineStyle.xlContinuous;
                    param.Weight = Excel.XlBorderWeight.xlThin;
                }
                else if ((Excel.XlBorderWeight)param.Weight == Excel.XlBorderWeight.xlThin)
                {
                    param.Weight = Excel.XlBorderWeight.xlMedium;
                }
                else if ((Excel.XlBorderWeight)param.Weight == Excel.XlBorderWeight.xlMedium)
                {
                    param.LineStyle = Excel.XlLineStyle.xlDouble;
                }
                else
                {
                    param.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void LowerBorderRange__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null || selection?.Borders == null) return;

                Excel.Borders borders = selection.Borders;
                Excel.Border param = borders[Excel.XlBordersIndex.xlEdgeBottom];

                if ((Excel.XlLineStyle)param.LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                {
                    param.LineStyle = Excel.XlLineStyle.xlDot;
                }
                else if ((Excel.XlLineStyle)param.LineStyle == Excel.XlLineStyle.xlDot)
                {
                    param.LineStyle = Excel.XlLineStyle.xlContinuous;
                    param.Weight = Excel.XlBorderWeight.xlThin;
                }
                else if ((Excel.XlBorderWeight)param.Weight == Excel.XlBorderWeight.xlThin)
                {
                    param.Weight = Excel.XlBorderWeight.xlMedium;
                }
                else if ((Excel.XlBorderWeight)param.Weight == Excel.XlBorderWeight.xlMedium)
                {
                    param.LineStyle = Excel.XlLineStyle.xlDouble;
                }
                else
                {
                    param.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void RightBorderRange__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null || selection?.Borders == null) return;

                Excel.Borders borders = selection.Borders;
                Excel.Border param = borders[Excel.XlBordersIndex.xlEdgeRight];

                if ((Excel.XlLineStyle)param.LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                {
                    param.LineStyle = Excel.XlLineStyle.xlDot;
                }
                else if ((Excel.XlLineStyle)param.LineStyle == Excel.XlLineStyle.xlDot)
                {
                    param.LineStyle = Excel.XlLineStyle.xlContinuous;
                    param.Weight = Excel.XlBorderWeight.xlThin;
                }
                else if ((Excel.XlBorderWeight)param.Weight == Excel.XlBorderWeight.xlThin)
                {
                    param.Weight = Excel.XlBorderWeight.xlMedium;
                }
                else if ((Excel.XlBorderWeight)param.Weight == Excel.XlBorderWeight.xlMedium)
                {
                    param.LineStyle = Excel.XlLineStyle.xlDouble;
                }
                else
                {
                    param.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void LeftBorderRange__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null || selection?.Borders == null) return;

                Excel.Borders borders = selection.Borders;
                Excel.Border param = borders[Excel.XlBordersIndex.xlEdgeLeft];

                if ((Excel.XlLineStyle)param.LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                {
                    param.LineStyle = Excel.XlLineStyle.xlDot;
                }
                else if ((Excel.XlLineStyle)param.LineStyle == Excel.XlLineStyle.xlDot)
                {
                    param.LineStyle = Excel.XlLineStyle.xlContinuous;
                    param.Weight = Excel.XlBorderWeight.xlThin;
                }
                else if ((Excel.XlBorderWeight)param.Weight == Excel.XlBorderWeight.xlThin)
                {
                    param.Weight = Excel.XlBorderWeight.xlMedium;
                }
                else if ((Excel.XlBorderWeight)param.Weight == Excel.XlBorderWeight.xlMedium)
                {
                    param.LineStyle = Excel.XlLineStyle.xlDouble;
                }
                else
                {
                    param.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void OuterBorderRange__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null || selection?.Borders == null) return;

                Excel.Borders borders = selection.Borders;
                Excel.Border topBorder = borders[Excel.XlBordersIndex.xlEdgeTop];
                Excel.Border bottomBorder = borders[Excel.XlBordersIndex.xlEdgeBottom];
                Excel.Border rightBorder = borders[Excel.XlBordersIndex.xlEdgeRight];
                Excel.Border leftBorder = borders[Excel.XlBordersIndex.xlEdgeLeft];

                List<Excel.Border> selectionBorders = new List<Excel.Border> { topBorder, bottomBorder, rightBorder, leftBorder };


                if ((Excel.XlLineStyle)topBorder.LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                {
                    foreach (Excel.Border border in selectionBorders)
                    {
                        border.LineStyle = Excel.XlLineStyle.xlDot;
                    }
                }
                else if ((Excel.XlLineStyle)topBorder.LineStyle == Excel.XlLineStyle.xlDot)
                {
                    foreach (Excel.Border border in selectionBorders)
                    {
                        border.LineStyle = Excel.XlLineStyle.xlContinuous;
                        border.Weight = Excel.XlBorderWeight.xlThin;
                    }
                }
                else if ((Excel.XlBorderWeight)topBorder.Weight == Excel.XlBorderWeight.xlThin)
                {
                    foreach (Excel.Border border in selectionBorders)
                    {
                        border.Weight = Excel.XlBorderWeight.xlMedium;
                    }
                }
                else if ((Excel.XlBorderWeight)topBorder.Weight == Excel.XlBorderWeight.xlMedium)
                {
                    foreach (Excel.Border border in selectionBorders)
                    {
                        border.LineStyle = Excel.XlLineStyle.xlDouble;
                    }
                }
                else
                {
                    foreach (Excel.Border border in selectionBorders)
                    {
                        border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    }
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void AllBorderRange__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null || selection?.Borders == null) return;

                Excel.Borders borders = selection.Borders;
                Excel.Border topBorder = borders[Excel.XlBordersIndex.xlEdgeTop];

                if ((Excel.XlLineStyle)topBorder.LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                {
                    borders.LineStyle = Excel.XlLineStyle.xlDot;
                }
                else if ((Excel.XlLineStyle)topBorder.LineStyle == Excel.XlLineStyle.xlDot)
                {
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    borders.Weight = Excel.XlBorderWeight.xlThin;
                }
                else if ((Excel.XlBorderWeight)topBorder.Weight == Excel.XlBorderWeight.xlThin)
                {
                    borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
                else if ((Excel.XlBorderWeight)topBorder.Weight == Excel.XlBorderWeight.xlMedium)
                {
                    borders.LineStyle = Excel.XlLineStyle.xlDouble;
                }
                else
                {
                    borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void NoBorders__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null || selection?.Borders == null) return;

                Excel.Borders borders = selection.Borders;

                borders.Weight = Excel.XlBorderWeight.xlThin;
                borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ColumnWidth__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Columns == null || selection == null) return;

                double param = Math.Round(activeCell.Columns.ColumnWidth, 2);
                if (param == 1.89)
                {
                    param = 5;
                }
                else if (param < 5)
                {
                    param = 5;
                }
                else if (param == 5)
                {
                    param = 8.11;
                }
                else if (param <= 8.2 && param >= 8)
                {
                    param = 10;
                }
                else if (param < 8.11)
                {
                    param = 8.11;
                }
                else if (param >= 35)
                {
                    param = 1.89;
                }
                else
                {
                    param += 5;
                }
                selection.Columns.ColumnWidth = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ColumnAutoFit__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null) return;

                int maxCols = 500;

                if (selection.Columns.Count >= maxCols)
                {
                    DialogResult result = System.Windows.Forms.MessageBox.Show("WARNING! You have selected " + selection.Columns.Count.ToString("#,##0") + " columns.\n\nThis action could take a significant amount of time.\n\nDo you wish to continue?", "WARNING",
                                                               MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (result != DialogResult.Yes) { return; }
                }

                selection.Columns.AutoFit();

                //int count = 0;
                //foreach (Excel.Range column in selection.Columns)
                //{
                //    count += 1;
                //    if (count >= maxCellsSlow) return;
                //    column.AutoFit();
                //}
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void RowHeight__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Columns == null || selection == null) return;

                double param = Math.Round(activeCell.Rows.RowHeight, 2);
                if (param < 14.4)
                {
                    param = 14.4;
                }
                else if (param == 14.4)
                {
                    param = 15;
                }
                else if (param == 15)
                {
                    param = 20;
                }
                else if (param >= 50)
                {
                    param = 14.4;
                }
                else
                {
                    param += 10;
                }
                selection.Rows.RowHeight = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void RowAutoFit__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                int maxRows = 500;

                if (selection.Rows.Count >= maxRows)
                {
                    DialogResult result = System.Windows.Forms.MessageBox.Show("WARNING! You have selected " + selection.Rows.Count.ToString("#,##0") + " rows.\n\nThis action could take a significant amount of time.\n\nDo you wish to continue?", "WARNING",
                                                               MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (result != DialogResult.Yes) { return; }
                }

                selection.Rows.AutoFit();

                //int count = 0;
                //foreach (Excel.Range row in selection.Rows)
                //{
                //    count += 1;
                //    if (count >= maxCellsSlow) return;
                //    else if (count >= 20000) return;
                //    row.AutoFit();
                //}
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void IncreaseFontSize__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Font == null || selection == null) return;

                double? param = Math.Round(activeCell.Font.Size, 2);

                if (param == null) param = 10;
                if (param >= 48) param = 1;
                else if (param != Math.Round((double)param)) param = Math.Ceiling((double)param);
                else param += 1;
                selection.Font.Size = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        } // check
        private void DecreaseFontSize__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || activeCell?.Font == null || selection == null) return;

                double? param = Math.Round(activeCell.Font.Size, 2);

                if (param == null) param = 10;
                if (param <= 1) param = 48;
                else if (param != Math.Round((double)param)) param = Math.Floor((double)param);
                else param -= 1;
                selection.Font.Size = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        } // check
        private void HorizontalAlignCycle__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || selection == null) return;

                int? param = activeCell.HorizontalAlignment;

                if (param == null)
                {
                    selection.HorizontalAlignment = -4131;
                    return;
                }

                if (param == -4131)
                {
                    param = -4108;
                }
                else if (param == -4108)
                {
                    param = -4152;
                }
                else
                {
                    param = -4131;
                }
                selection.HorizontalAlignment = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void ToggleCenterOverSelection__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || selection == null) return;

                int? param = activeCell.HorizontalAlignment;

                if (param == 7)
                {
                    param = 1;
                }
                else
                {
                    param = 7;
                }
                selection.HorizontalAlignment = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void MergeCells__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || selection == null) return;

                bool? param = activeCell.MergeCells;

                if ((bool)param)
                {
                    selection.UnMerge();
                }
                else
                {
                    selection.Merge();
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void WrapText__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || selection == null) return;

                bool? param = activeCell.WrapText;

                if ((bool)param)
                {
                    selection.WrapText = false;
                }
                else
                {
                    selection.WrapText = true;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void VerticalAlignCycle__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || selection == null) return;

                int? param = activeCell.VerticalAlignment;

                if (param == -4107)
                {
                    param = -4108;
                }
                else if (param == -4108)
                {
                    param = -4160;
                }
                else
                {
                    param = -4107;
                }
                selection.VerticalAlignment = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void IncreaseIndent__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || selection == null) return;

                int? param = activeCell.IndentLevel;

                if (param == null)
                {
                    param = 0;
                }

                if (param > 15)
                {
                    param = 0;
                }
                else
                {
                    param += 1;
                }
                selection.IndentLevel = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void DecreaseIndent__()
        {
            Excel.Window excelApp = null;
            Excel.Range selection = null;
            Excel.Range activeCell = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application.ActiveWindow;
                selection = excelApp.Selection;
                activeCell = excelApp.ActiveCell;

                if (activeCell == null || selection == null) return;

                int? param = activeCell.IndentLevel;

                if (param == null)
                {
                    param = 0;
                }

                if (param > 15)
                {
                    param = 0;
                }
                else
                {
                    param -= 1;
                }
                selection.IndentLevel = param;
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (activeCell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeCell);
                }
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void IfError__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (TooManyCells(maxCellsSlow)) return;
                int count = 0;

                string err = this.FunctionData["IfError"][0];

                foreach (Excel.Range cell in selection.Cells)
                {
                    // stop after X cells
                    count += 1;
                    if (count >= maxCellsSlow) return;

                    if (cell.Value == null)
                    {
                        continue;
                    }

                    // formula
                    string form = cell.Formula;
                    if (string.IsNullOrWhiteSpace(form) || form[0] != '=') // Check if formula starts with '='
                    {
                        continue;
                    }

                    // add iferror if doesn't start with
                    if (!(form.StartsWith("=IFERROR(") && form.EndsWith(")")))
                    {
                        form = "=IFERROR(" + form.Substring(1) + ",\"" + err + "\")";
                        cell.Formula = form;
                        continue;
                    }

                    // find comma arg separator
                    int commaIndex = 0;
                    int endParenIndex = 1;
                    Stack<char> stack = new Stack<char>();
                    for (int i = 0; i < form.Length; i++)
                    {
                        if (form[i] == '(') stack.Push('(');
                        else if (form[i] == ')')
                        {
                            stack.Pop();
                            if (stack.Count == 0) endParenIndex = i;
                        }
                        if (form[i] == ',' && stack.Count == 1) commaIndex = i;
                    }

                    string oldErr = form.Substring(commaIndex + 1, endParenIndex - commaIndex - 1);

                    // check parentheses
                    // copy formula, check if formula would cause error due to unmatched parentheses
                    stack = new Stack<char>();
                    //Stack<char> stack = new Stack<char>();
                    bool unmatched = false;
                    string form1 = "=" + form.Substring(9, form.Length - oldErr.Length - 11);

                    for (int i = 0; i < form1.Length; i++)
                    {
                        if (form1[i] == '(')
                        {
                            stack.Push('(');
                        }
                        else if (form1[i] == ')')
                        {
                            if (stack.Count == 0)
                            {
                                unmatched = true;
                                break;
                            }
                            else
                            {
                                stack.Pop();
                            }
                        }
                    }

                    // check for unmatched opening parentheses
                    // if unmatched == true or there are remaining parentheses in the stack
                    unmatched = unmatched || stack.Count > 0;

                    // if formula would be unmatched, make negative. Otherwise, make positive
                    if (unmatched)
                    {
                        form = "=IFERROR(" + form.Substring(1) + ",\"" + err + "\")";
                    }
                    else
                    {
                        form = "=" + form.Substring(9, form.Length - oldErr.Length - 11);
                    }
                    cell.Formula = form;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void CleanCells__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                string sheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
                if (sheetName.Contains(" "))
                {
                    sheetName = "'" + sheetName + "'!";
                }
                else
                {
                    sheetName = sheetName + "!";
                }

                if (TooManyCells(maxCellsSlow)) return;
                int count = 0;

                foreach (Excel.Range cell in selection.Cells)
                {
                    // stop after X cells
                    count += 1;
                    if (count >= maxCellsSlow) return;

                    string form = cell.Formula;
                    if (string.IsNullOrWhiteSpace(form) || form[0] != '=') // Check if formula starts with '='
                    {
                        continue;
                    }

                    Regex regex = new Regex("'([^']*)'");

                    string placeholder = "{=+=}";

                    List<Match> matches = regex.Matches(form).Cast<Match>().ToList();
                    foreach (Match match in matches)
                    {
                        form = form.Replace(match.Value, placeholder);
                    }

                    form = form.Replace(" ", "");

                    foreach (Match match in matches)
                    {
                        int placeholderIndex = form.IndexOf(placeholder, StringComparison.Ordinal);
                        if (placeholderIndex >= 0)
                        {
                            form = form.Remove(placeholderIndex, placeholder.Length);
                            form = form.Insert(placeholderIndex, match.Value);
                        }
                    }

                    if (form.Contains(sheetName))
                    {
                        form = form.Replace(sheetName, "");
                    }
                    cell.Formula = form;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }
        private void FlattenSelection__()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (TooManyCells(maxCellsSlow)) return;
                int count = 0;

                foreach (Excel.Range cell in selection.Cells)
                {
                    // stop after X cells
                    count += 1;
                    if (count >= maxCellsSlow) return;

                    if (cell.Value == null)
                    {
                        continue;
                    }

                    // formula
                    string form = cell.Formula;
                    if (string.IsNullOrWhiteSpace(form)) // Check if formula starts with '='
                    {
                        continue;
                    }

                    cell.Formula = cell.Value;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }

        #endregion

        #region Functions
        [Obfuscation(Exclude = true)]
        private void BlankFunction() { }

        [Obfuscation(Exclude = true)]
        public void ToggleCellSign() { ToggleCellSign__(); }

        [Obfuscation(Exclude = true)]
        private void ToggleFontColor() { ToggleFontColor__(); }

        [Obfuscation(Exclude = true)]
        private void CycleFontColor1() { CycleFontColor1__(); }

        [Obfuscation(Exclude = true)]
        private void CycleFontColor2() { CycleFontColor2__(); }

        [Obfuscation(Exclude = true)]
        private void AutoColorCells() { AutoColorCells__(); }

        [Obfuscation(Exclude = true)]
        private void CycleFillColor2() { CycleFillColor2__(); }

        [Obfuscation(Exclude = true)]
        private void CycleFillColor1() { CycleFillColor1__(); }

        [Obfuscation(Exclude = true)]
        public void IncreaseDecimal() { IncreaseDecimal__(); }

        [Obfuscation(Exclude = true)]
        public void DecreaseDecimal() { DecreaseDecimal__(); }

        [Obfuscation(Exclude = true)]
        public void ShiftDecimalLeft() { ShiftDecimalLeft__(); }

        [Obfuscation(Exclude = true)]
        public void ShiftDecimalRight() { ShiftDecimalRight__(); }

        [Obfuscation(Exclude = true)]
        public void ToggleGeneralNumberFormats() { ToggleGeneralNumberFormats__(); }

        [Obfuscation(Exclude = true)]
        public void CycleDateFormats() { CycleDateFormats__(); }

        [Obfuscation(Exclude = true)]
        public void CycleCurrency() { CycleCurrency__(); }

        [Obfuscation(Exclude = true)]
        public void CycleForeignCurrency() { CycleForeignCurrency__(); }

        [Obfuscation(Exclude = true)]
        public void CyclePercent() { CyclePercent__(); }

        [Obfuscation(Exclude = true)]
        private void ToggleBPS() { ToggleBPS__(); }

        [Obfuscation(Exclude = true)]
        public void ToggleMultiple() { ToggleMultiple__(); }

        [Obfuscation(Exclude = true)]
        public void ToggleBinary() { ToggleBinary__(); }

        [Obfuscation(Exclude = true)]
        public void CycleOtherNumberFormats1() { CycleOtherNumberFormats1__(); }

        [Obfuscation(Exclude = true)]
        public void CycleOtherNumberFormats2() { CycleOtherNumberFormats2__(); }

        [Obfuscation(Exclude = true)]
        public void CycleOtherNumberFormats3() { CycleOtherNumberFormats3__(); }

        [Obfuscation(Exclude = true)]
        private void UpperBorderRange() { UpperBorderRange__(); }

        [Obfuscation(Exclude = true)]
        private void LowerBorderRange() { LowerBorderRange__(); }

        [Obfuscation(Exclude = true)]
        private void RightBorderRange() { RightBorderRange__(); }

        [Obfuscation(Exclude = true)]
        private void LeftBorderRange() { LeftBorderRange__(); }

        [Obfuscation(Exclude = true)]
        private void OuterBorderRange() { OuterBorderRange__(); }

        [Obfuscation(Exclude = true)]
        private void AllBorderRange() { AllBorderRange__(); }

        [Obfuscation(Exclude = true)]
        private void NoBorders() { NoBorders__(); }

        [Obfuscation(Exclude = true)]
        private void ColumnWidth() { ColumnWidth__(); }

        [Obfuscation(Exclude = true)]
        private void ColumnAutoFit() { ColumnAutoFit__(); }

        [Obfuscation(Exclude = true)]
        private void RowHeight() { RowHeight__(); }

        [Obfuscation(Exclude = true)]
        private void RowAutoFit() { RowAutoFit__(); }

        [Obfuscation(Exclude = true)]
        private void IncreaseFontSize() { IncreaseFontSize__(); }

        [Obfuscation(Exclude = true)]
        private void DecreaseFontSize() { DecreaseFontSize__(); }

        [Obfuscation(Exclude = true)]
        private void HorizontalAlignCycle() { HorizontalAlignCycle__(); }

        [Obfuscation(Exclude = true)]
        private void ToggleCenterOverSelection() { ToggleCenterOverSelection__(); }

        [Obfuscation(Exclude = true)]
        private void MergeCells() { MergeCells__(); }

        [Obfuscation(Exclude = true)]
        private void WrapText() { WrapText__(); }

        [Obfuscation(Exclude = true)]
        private void VerticalAlignCycle() { VerticalAlignCycle__(); }

        [Obfuscation(Exclude = true)]
        private void IncreaseIndent() { IncreaseIndent__(); }

        [Obfuscation(Exclude = true)]
        private void DecreaseIndent() { DecreaseIndent__(); }

        [Obfuscation(Exclude = true)]
        public void IfError() { IfError__(); }

        [Obfuscation(Exclude = true)]
        public void CleanCells() { CleanCells__(); }

        [Obfuscation(Exclude = true)]
        public void FlattenSelection() { FlattenSelection__(); }

        #endregion

        #region Other
        private bool TooManyCells(int maxCells)
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection == null) return true;

                if (selection.Count >= maxCells)
                {
                    DialogResult result = System.Windows.Forms.MessageBox.Show("WARNING! You have selected " + selection.Count.ToString("#,##0") + " cells.\n\nThis action could take a significant amount of time.\n\nDo you wish to continue?", "WARNING",
                                                               MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (result != DialogResult.Yes) { return true; }
                    return false;
                }
                return false;
            }
            catch (COMException) { return false; }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.Message);
                return false;
            }
            finally
            {
                if (selection != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                }
            }
        }

        private int RGB(int r, int g, int b)
        {
            return (r & 0xFF) | ((g & 0xFF) << 8) | ((b & 0xFF) << 16);
        }
        private static bool IsExcelActive()
        {
            IntPtr handle = GetForegroundWindow();
            Excel.Window excelApp = Globals.ThisAddIn.Application.ActiveWindow;

            // Check if the active window belongs to Excel.
            if (handle == new IntPtr(excelApp.Hwnd))
            {
                // Check if the active window belongs to an Excel workbook.
                Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook; // was excelApp
                if (activeWorkbook != null)
                {
                    // Check if the active window belongs to an Excel worksheet.
                    Excel.Worksheet activeWorksheet = excelApp.ActiveSheet;
                    if (activeWorksheet != null)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        private bool? ShowConfirmationWindow(string confirmationMessage)
        {
            ConfirmationForm confirmationForm = new ConfirmationForm(confirmationMessage);

            // Set the owner of the WPF window to the Excel application.
            WindowInteropHelper windowInteropHelper = new WindowInteropHelper(confirmationForm);
            windowInteropHelper.Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd);

            return confirmationForm.ShowDialog();
        }
        public bool? ShowErrorMessage(string errorMessage)
        {
            ErrorForm errorForm = new ErrorForm(errorMessage);

            // Set the owner of the WPF window to the Excel application.
            WindowInteropHelper windowInteropHelper = new WindowInteropHelper(errorForm);
            windowInteropHelper.Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd);

            return errorForm.ShowDialog();
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
