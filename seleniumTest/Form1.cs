using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace seleniumTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void CallSelenium()
        {
            MainProcess mp = new MainProcess();
            mp.MyForm = this;
            mp.MyRichTextBox = this.statusRTB;
            mp.FileName = this.FileTxtBox.Text;
            mp.Key = this.textBox1.Text;

            mp.Run();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!(String.IsNullOrEmpty(FileTxtBox.Text) || String.IsNullOrEmpty(textBox1.Text)))
            {
                Thread t = new Thread(CallSelenium);
                t.Name = FileTxtBox.Text;
                t.Start();
            }
            else
                MessageBox.Show("Please fill the input fields!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                FileTxtBox.Text = openFileDialog1.FileName;
            }
        }
    }
}
