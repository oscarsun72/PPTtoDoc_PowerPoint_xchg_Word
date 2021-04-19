using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Doc_PPt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new TableDocOps().splitTableByEachRowTitleed字源圖片();
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text= Clipboard.GetText();
        }
    }
}
