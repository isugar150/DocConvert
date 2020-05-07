using System;
using System.Windows.Forms;

namespace DocConvert_Server
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
    }
}
