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
        public string cob;
        public List<string> ayy = new List<string>();
        public frmChangequilty()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txt = this.textBox1.Text;
            cob = this.dianpu_cob.SelectedItem.ToString();
            if (txt != null && txt != "")
                this.Close();
            else
                MessageBox.Show("不能为空值");



        }

        private void frmChangequilty_Load(object sender, EventArgs e)
        {
          //  frm调拨系统 ff = new frm调拨系统();
            this.dianpu_cob.DataSource = ayy;
        }
    }
}
