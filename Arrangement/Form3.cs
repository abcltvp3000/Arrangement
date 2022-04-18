using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;
using System.Diagnostics;
using System.Text;

namespace Arrangement
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            this.Text = "Tiến trình";
        }

        public void delay(int milisec)
        {
            Thread.Sleep(milisec);
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
