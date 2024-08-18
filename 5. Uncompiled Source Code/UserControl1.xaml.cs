using System.Windows;

namespace GES
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class ConfirmationForm : System.Windows.Window
    {
        public ConfirmationForm(string confirmationMessage)
        {
            InitializeComponent();
            ConfirmLabel.Content = confirmationMessage;
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true; // sets the result and closes the window
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false; // sets the result and closes the window
        }
    }
}