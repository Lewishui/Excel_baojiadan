using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QCAuto
{
    public partial class frmChangequilty : Form
    {

        public string txt;

        public frmChangequilty()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txt = this.textBox1.Text;

            if (txt != null && txt != "")
                this.Close();
            else
                MessageBox.Show("不能为空值");



        }
    }
}
