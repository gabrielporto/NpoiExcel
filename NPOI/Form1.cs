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
            OpenFileDialog fd = new OpenFileDialog();
            if ((fd.ShowDialog() == System.Windows.Forms.DialogResult.OK))
            {
                Cursor.Current = Cursors.WaitCursor;
                dataGridView1.DataSource = excel.GetDataTable(fd.FileName);
            }

        }
    }
}
