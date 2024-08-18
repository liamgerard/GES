using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace GES
{
    public partial class UserControl1 : UserControl
    {

        private ElementHost elementHost;
        private CutsForm CutsForm;
        public UserControl1()
        {
            InitializeComponent();

            CutsForm = new CutsForm();
            elementHost = new ElementHost { Dock = DockStyle.Fill, Child = CutsForm };
            this.Controls.Add(elementHost);
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

        }
    }
}
