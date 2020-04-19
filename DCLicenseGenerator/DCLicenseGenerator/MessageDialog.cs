using System;
using System.Windows.Forms;

namespace DCLicenseGenerator
{
    public partial class MessageDialog : Form
    {
        public MessageDialog(string Title, string label, string textBox)
        {
            InitializeComponent();
            this.Text = Title;
            label1.Text = label;
            textBox1.Text = textBox;
        }

        private void MessageDialog_Load(object sender, EventArgs e)
        {

        }

        private void MessageDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
