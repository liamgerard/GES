using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace GES
{
    public partial class UserControl3 : UserControl
    {

        private ElementHost elementHost;
        private DisabledKeysForm disKeysForm;
        public UserControl3()
        {
            InitializeComponent();

            disKeysForm = new DisabledKeysForm();
            elementHost = new ElementHost { Dock = DockStyle.Fill, Child = disKeysForm };
            this.Controls.Add(elementHost);
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

        }
    }
}
