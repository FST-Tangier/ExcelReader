using System;
using System.Windows.Forms;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
           grd.DataSource =  Helper.Read("registre.xls");
        }
    }
}
