using System;
using System.Windows;

namespace GES
{
    public partial class ErrorForm : Window
    {
        public ErrorForm(string errorMessage)
        {
            InitializeComponent();
            ErrorText.Text = errorMessage;

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true; // sets the result and closes the window
        }
    }
}
