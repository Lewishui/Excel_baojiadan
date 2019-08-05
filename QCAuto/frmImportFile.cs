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
    public partial class frmImportFile : Form
    {
        public string path;
        public string path2;
        public frmImportFile()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog tbox = new OpenFileDialog();
            tbox.Multiselect = false;
            tbox.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            if (tbox.ShowDialog() == DialogResult.OK)
            {
                path = tbox.FileName;
                textBox1.Text = tbox.FileName;
            }
            if (path == null || path == "")
                return;
        }

        private void openFileBtton_Click(object sender, EventArgs e)
        {
            OpenFileDialog tbox = new OpenFileDialog();
            tbox.Multiselect = false;
            tbox.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            if (tbox.ShowDialog() == DialogResult.OK)
            {
                path2 = tbox.FileName;
                pathTextBox.Text = tbox.FileName;
            }
            if (path2 == null || path2 == "")
                return;
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            if (path2 == null || path2 == "" || path == null || path == "")
            {
                MessageBox.Show("请选择文件");
                return;
            }
            else

                this.Close();

        }
    }
}
