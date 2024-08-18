using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace GES
{
    public partial class UserControl2 : UserControl
    {
        private ElementHost elementHost;
        private FormatForm FormatForm;
        public UserControl2()
        {
            InitializeComponent();

            FormatForm = new FormatForm();
            elementHost = new ElementHost() { Dock = DockStyle.Fill };
            elementHost.Child = FormatForm;
            this.Controls.Add(elementHost);
        }

        private void UserControl2_Load(object sender, System.EventArgs e)
        {

        }
    }
}
