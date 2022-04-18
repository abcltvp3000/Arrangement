using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Arrangement
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.Text = "Danh sách phân công ngẫu nhiên";
        }

        void expExcel(DataGridView grid, string title = "")
        {
            SaveFileDialog file = new SaveFileDialog();
            file.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            file.Title = title;
            file.ShowDialog();
            string path = file.FileName.ToString();
            if (path == "")
            {
                return;
            }

            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.Visible = true;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = title;
            worksheet.Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            worksheet.Cells.Borders.Weight = 2d;
            worksheet.Rows.AutoFit();
            for (int i = 1; i < grid.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i].NumberFormat = "@";
                worksheet.Cells[1, i].Font.Bold = true;
                worksheet.Cells[1, i].Font.Size = 12;
                worksheet.Cells[1, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                worksheet.Cells[1, i].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[1, i].Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[1, i] = grid.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < grid.Rows.Count - 1; i++)
            {
                for (int j = 0; j < grid.Columns.Count; j++)
                {
                    if (j == 0) worksheet.Cells[i + 2, j + 1].NumberFormat = "@";
                    worksheet.Cells[i + 2, j + 1] = grid.Rows[i].Cells[j].Value.ToString();
                }
            }
            workbook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.Quit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            expExcel(arrangeGrid, "DS phân công theo đơn vị");
        }
    }
}
