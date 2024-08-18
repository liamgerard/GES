using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using Xceed.Wpf.Toolkit;
using Excel = Microsoft.Office.Interop.Excel;

namespace GES
{
    /// <summary>
    /// Interaction logic for UserControl6.xaml
    /// </summary>
    public partial class ColorFormPopup : System.Windows.Window
    {
        public Color SelectedColor
        {
            get { return (Color)_colorCanvas.SelectedColor; }
            set { _colorCanvas.SelectedColor = value; }
        }
        public ObservableCollection<Xceed.Wpf.Toolkit.ColorItem> ColorList;
        public ColorFormPopup()
        {
            InitializeComponent();
            PopulateColorListThemeColors();
        }
        private void PopulateColorListThemeColors()
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                ObservableCollection<Xceed.Wpf.Toolkit.ColorItem> ColorList = new ObservableCollection<Xceed.Wpf.Toolkit.ColorItem>();
                ThemeColorScheme themeColors = workbook.Theme.ThemeColorScheme;
                int count = 1;

                foreach (ThemeColor themeColor in themeColors)
                {
                    if (count > themeColors.Count) break;

                    Color color = ConvertThemeColorToMediaColor((XlThemeColor)count);
                    ColorList.Add(new ColorItem(color, color.ToString()));
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (workbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
            }
        }
        public void saveButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
        public Color ConvertThemeColorToMediaColor(Excel.XlThemeColor themeColor)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            ThemeColorScheme themeColors = excelApp.ActiveWorkbook.Theme.ThemeColorScheme;

            // Get the theme color as an RGB
            int oldColor = System.Convert.ToInt32(themeColor);

            // Convert the RGB OLE color to a System.Drawing.Color
            System.Drawing.Color drawingColor = System.Drawing.ColorTranslator.FromOle(oldColor);

            // Convert System.Drawing.Color to System.Windows.Media.Color
            Color mediaColor = Color.FromArgb(drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B);

            return mediaColor;
        }
    }
}
