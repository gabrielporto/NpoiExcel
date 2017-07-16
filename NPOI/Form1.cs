using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//add
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using System.Data.OleDb;
using System.Collections;

namespace NPOI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Excel excel = new Excel();

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if ((ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK))
            {
                Cursor.Current = Cursors.WaitCursor;
                comboBox1.DataSource = excel.GetTabName(ofd.FileName);
                Cursor.Current = Cursors.Default;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if ((sfd.ShowDialog()) == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                excel.SalvarDataTable((DataTable)dataGridView1.DataSource, sfd.FileName);
                Cursor.Current = Cursors.Default;
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = excel.GetDataTable(null,comboBox1.SelectedValue.ToString());
        }
    }
}
