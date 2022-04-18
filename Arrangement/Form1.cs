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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            this.Text = "Phần mềm phân công coi thi THPT - QG năm học 2021-2022";
        }

        string showExcelDialog(string title = "")
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            file.Title = title;
            file.ShowDialog();
            string path = file.FileName.ToString();
            return path;
        }

        string compressPath(string path)
        {
            if (path.Length > 35)
            {
                return "..." + path.Substring(path.Length - 32);
            }
            return path;
        }

        string statImpPath = "";
        bool statImpType = false;
        bool is_Calc = false;
        private void statImpBtn_Click(object sender, EventArgs e)
        {
            statImpType = false;
            statImpPath = showExcelDialog("DS đề xuất theo đơn vị");
            statImpName.Text = (statImpPath == "" ? "Chưa có file nào được chọn": compressPath(statImpPath));
            statImpName2.Text = "Chưa có file nào được chọn";
        }

        string objToString(object o)
        {
            string str = o?.ToString() ?? "";
            return str;
        }

        int JobType(string str)
        {
            int jobType = -1;
            switch (str)
            {
                case "Trưởng điểm":
                    jobType = 0;
                    break;
                case "Phó Trưởng điểm":
                    jobType = 1;
                    break;
                case "Thư ký":
                    jobType = 2;
                    break;
                case "Cán bộ coi thi":
                    jobType = 3;
                    break;
                case "Cán bộ giám sát":
                    jobType = 4;
                    break;
                default:
                    break;
            }
            return jobType;
        }

        string JobStr(int x)
        {
            string str = "";
            switch (x)
            {
                case 0:
                    str = "Trưởng điểm";
                    break;
                case 1:
                    str = "Phó Trưởng điểm";
                    break;
                case 2:
                    str = "Thư ký";
                    break;
                case 3:
                    str = "Cán bộ coi thi";
                    break;
                case 4:
                    str = "Cán bộ giám sát";
                    break;
                default:
                    break;
            }
            return str;
        }

        Dictionary<string, List<string>[]> statNames = new Dictionary<string, List<string>[]>();
        public void impStatGrid(string path)
        {
            var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            var tables = result.Tables.Cast<DataTable>();
            if (!statImpType) {
                foreach (DataTable table in tables)
                {
                    if (table.Columns.Count != statGrid.Columns.Count)
                    {
                        continue;
                    }
                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        statGrid.Rows.Add(table.Rows[i].ItemArray);
                    }
                }
            }
            else
            {
                Dictionary<string, string> schoolNames = new Dictionary<string, string>();
                Dictionary<string, int[]> statSchool = new Dictionary<string, int[]>();
                foreach (DataTable table in tables)
                {
                    if (table.Columns.Count < 12)
                    {
                        continue;
                    }
                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        string str = objToString(table.Rows[i][1]);
                        schoolNames[str] = objToString(table.Rows[i][2]);
                    }
                }
                foreach (var school in schoolNames)
                {
                    statSchool[school.Key] = new int[5];
                }
                int count = statFullGrid.Rows.Count - 1;
                foreach (DataTable table in tables)
                {
                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        string str = objToString(table.Rows[i][12]);
                        string code = objToString(table.Rows[i][1]);
                        int jobType = JobType(str);
                        if (jobType != -1)
                        {
                            string fullName = objToString(table.Rows[i][3]);
                            statSchool[code][jobType]++;
                            statFullGrid.Rows.Add();
                            statFullGrid.Rows[count].Cells[0].Value = code;
                            statFullGrid.Rows[count].Cells[1].Value = schoolNames[code];
                            statFullGrid.Rows[count].Cells[2].Value = fullName;
                            statFullGrid.Rows[count].Cells[3].Value = str;
                            count++;
                        }
                    }
                }

                count = statGrid.Rows.Count - 1;
                foreach (var school in statSchool)
                {
                    int sum = 0;
                    for (int i = 0; i < 5; i++) sum += school.Value[i];
                    if (sum == 0) continue;
                    statGrid.Rows.Add();
                    statGrid.Rows[count].Cells[0].Value = school.Key;
                    statGrid.Rows[count].Cells[1].Value = schoolNames[school.Key];
                    for (int j = 0; j < 5; j++)
                    {
                        statGrid.Rows[count].Cells[2 + j].Value = school.Value[j];
                    }
                    count++;
                }
            }
            stream.Close();
        }

        private void statAcpImpBtn_Click(object sender, EventArgs e)
        {
            if (statImpPath != "")
            {
                is_Calc = false;
                impStatGrid(statImpPath);
            }
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

        private void statExpBtn_Click(object sender, EventArgs e)
        {
            expExcel(statGrid, "DS đề xuất theo đơn vị");
        }

        string needImpPath = "";
        //bool needImpType = false;
        private void needImpBtn_Click(object sender, EventArgs e)
        {
            needImpPath = showExcelDialog("DS thống kê theo điểm thi");
            needImpName.Text = (needImpPath == "" ? "Chưa có file nào được chọn" : compressPath(needImpPath));
        }

        public void impNeedGrid(string path)
        {
            var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            var tables = result.Tables.Cast<DataTable>();
            foreach (DataTable table in tables)
            {
                if (table.Columns.Count != needGrid.Columns.Count)
                {
                    continue;
                }
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    needGrid.Rows.Add(table.Rows[i].ItemArray);
                }
            }
            stream.Close();
        }

        private void needAcpImpBtn_Click(object sender, EventArgs e)
        {
            if (needImpPath != "")
            {
                is_Calc = false;
                impNeedGrid(needImpPath);
            }
        }

        private void needExpBtn_Click(object sender, EventArgs e)
        {
            expExcel(needGrid, "DS thống kê theo điểm thi");
        }

        int[][][] result = new int[0][][];
        Dictionary<string, string> schoolNames = new Dictionary<string, string>();
        Dictionary<string, int> schools = new Dictionary<string, int>();
        //Dictionary<string, List<Tuple<string, int, int, int>>> Result = new Dictionary<string, List<Tuple<string, int, int, int>>>();
        //Dictionary<string, List<Tuple<string, int, int, int>>> Result2 = new Dictionary<string, List<Tuple<string, int, int, int>>>();
        private void resultCalcBtn_Click(object sender, EventArgs e)
        {
            is_Calc = true;
            statNames = new Dictionary<string, List<string>[]>();
            schoolNames = new Dictionary<string, string>();
            schools = new Dictionary<string, int>();

            var statSchool = new Dictionary<string, int[]>();
            int[] lacks = new int[5];
            for (int i = 0; i < statGrid.Rows.Count - 1; i++)
            {
                int[] arr2 = new int[5];
                for (int j = 0; j < 5; j++)
                {
                    string tmp = statGrid.Rows[i].Cells[j + 2].Value?.ToString() ?? "";
                    try
                    {
                        arr2[j] = Convert.ToInt32(tmp);
                    }
                    catch
                    {
                        arr2[j] = 0;
                    }
                }
                string schoolCode = "";
                /*var cell = statGrid.Rows[i].Cells[0];
                if (cell is not null)
                { 
                    schoolCode = cell.Value.ToString();
                }*/
                schoolCode = statGrid.Rows[i].Cells[0].Value?.ToString() ?? "";
                if (schoolCode == "") continue;
                statNames[schoolCode] = new List<string>[5];
                for (int j = 0; j < 5; j++)
                {
                    statNames[schoolCode][j] = new List<string>();
                }
                statSchool[schoolCode] = arr2;
                schools[schoolCode] = 0;
                string schoolName = statGrid.Rows[i].Cells[1].Value?.ToString() ?? "";
                schoolNames[schoolCode] = schoolName;
            }
            if (statImpType)
            {
                for (int i = 0; i < statFullGrid.Rows.Count - 1; i++)
                {
                    string schoolCode = statFullGrid.Rows[i].Cells[0].Value?.ToString() ?? "";
                    if (schoolCode == "") continue;
                    string name = statFullGrid.Rows[i].Cells[2].Value?.ToString() ?? "";
                    if (name == "") continue;
                    int jobType = JobType(statFullGrid.Rows[i].Cells[3].Value?.ToString() ?? "");
                    if (jobType != -1) statNames[schoolCode][jobType].Add(name);
                }
            }

            foreach (var school in statSchool)
            {
                for (int j = 0; j < 5; j++)
                {
                    lacks[j] -= school.Value[j];
                }
            }

            Dictionary<string, int[]> needSchool = new Dictionary<string, int[]>();
            for (int i = 0; i < needGrid.Rows.Count - 1; i++)
            {
                int[] arr2 = new int[5];
                for (int j = 0; j < 5; j++)
                {
                    string tmp = needGrid.Rows[i].Cells[j + 2].Value?.ToString() ?? "";
                    try
                    {
                        arr2[j] = Convert.ToInt32(tmp);
                    }
                    catch (Exception)
                    {
                        arr2[j] = 0;
                    }
                }
                string schoolCode = "";
                /*var cell = needGrid.Rows[i].Cells[0];
                if (cell is not null)
                { 
                    schoolCode = cell.Value.ToString();
                }*/
                schoolCode = needGrid.Rows[i].Cells[0].Value?.ToString() ?? "";
                if (schoolCode == "") continue;

                needSchool[schoolCode] = arr2;
                schools[schoolCode] = 0;
                string schoolName = needGrid.Rows[i].Cells[1].Value?.ToString() ?? "";
                schoolNames[schoolCode] = schoolName;
            }

            int count = 0;
            List<string> lSchools = new List<string>(), lNameSchools = new List<string>();
            foreach (KeyValuePair<string, int> pair in schools)
            {
                lSchools.Add(pair.Key);
                lNameSchools.Add(schoolNames[pair.Key]);
                schools[pair.Key] = count++;
            }

            List<Tuple<int, int>>[] needList = new List<Tuple<int, int>>[5];
            for (int j = 0; j < 5; j++)
            {
                needList[j] = new List<Tuple<int, int>>();
            }
            foreach (var school in needSchool)
            {
                for (int j = 0; j < 5; j++)
                {
                    needList[j].Add(new Tuple<int, int>(school.Value[j], schools[school.Key]));
                    lacks[j] += school.Value[j];
                }
            }

            int[][] Lacks = new int[5][];
            for (int j = 0; j < 5; j++)
            {
                Lacks[j] = new int[schools.Count];
                needList[j].Sort(); needList[j].Reverse();
                for (int i = 0; i < needList[j].Count; i++)
                {
                    (var x, var y) = needList[j][i];
                    Lacks[j][y] = (lacks[j] <= 0 ? 0: lacks[j] / needSchool.Count + Convert.ToInt32(i < (lacks[j] % needSchool.Count)));
                }
            }

            List<Tuple<int, int>>[] queue = new List<Tuple<int, int>>[5];
            for (int j = 0; j < 5; j++)
            {
                queue[j] = new List<Tuple<int, int>>();
                foreach (KeyValuePair<string, int[]> pair in statSchool)
                {
                    queue[j].Add(new Tuple<int, int>(schools[pair.Key], pair.Value[j]));
                }
                queue[j].Shuffle();
                queue[j].Shuffle();
            }

            //int[][][] result = new int[schools.Count][][];
            result = new int[schools.Count][][];
            for (int i = 0; i < schools.Count; i++)
            {
                result[i] = new int[schools.Count][];
                for (int j = 0; j < schools.Count; j++)
                {
                    result[i][j] = new int[5];
                }
            }

            count = 0;
            foreach (KeyValuePair<string, int[]> pair in needSchool)
            {
                int g = schools[pair.Key];
                for (int j = 0; j < 5; j++)
                {
                    int z = pair.Value[j] - Lacks[j][g]; // (lacks[j] <= 0 ? 0: lacks[j] / needSchool.Count + Convert.ToInt32(count < (lacks[j] % needSchool.Count)));
                    while (queue[j].Count > 0 && z > 0)
                    {
                        (var x, var y) = queue[j][queue[j].Count - 1];
                        int t = Math.Min(z, y);
                        z -= t;
                        y -= t;
                        result[x][g][j] += t;
                        if (y > 0)
                        {
                            queue[j][queue[j].Count - 1] = new Tuple<int, int>(x, y);
                        }
                        else queue[j].RemoveAt(queue[j].Count - 1);
                    }
                }
                count++;
            }

            /*resultGrid.Rows.Clear();
            resultGrid.Columns.Clear();
            DataGridViewColumn newCol = new DataGridViewTextBoxColumn();
            newCol.HeaderText = "Mã trường";
            newCol.Name = "resSchoolCode";
            newCol.Visible = true;
            newCol.Width = 100;
            resultGrid.Columns.Add(newCol);
            for (int i = 0; i < lSchools.Count; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    DataGridViewColumn newColj = new DataGridViewTextBoxColumn();
                    newColj.HeaderText = (j == 4 ? lSchools[i]: "");
                    newColj.Name = "resT" + Convert.ToString(i * 5 + j);
                    newColj.Visible = true;
                    newColj.Width = 30;
                    resultGrid.Columns.Add(newColj);
                }
            }
            resultGrid.Rows.Add();
            resultGrid.Rows[0].Cells[0].Value = "Phân loại";
            for (int i = 0; i < lSchools.Count; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    resultGrid.Rows[0].Cells[1 + i * 5 + j].Value = Convert.ToString(1 + j);
                }
            }
            for (int i = 0; i < lSchools.Count; i++)
            {
                resultGrid.Rows.Add();
                resultGrid.Rows[1 + i].Cells[0].Value = lSchools[i];
                for (int j = 0; j < lSchools.Count; j++)
                {
                    for (int k = 0; k < 5; k++)
                    {
                        resultGrid.Rows[1 + i].Cells[1 + j * 5 + k].Value = result[i][j][k];
                    }
                }
            }*/

            resultGrid.Rows.Clear();
            count = 0;
            for (int i = 0; i < lSchools.Count; i++)
            {
                int gr_cnt = 0;
                for (int j = 0; j < lSchools.Count; j++)
                {
                    if (i == j) continue;
                    int sum = 0;
                    for (int k = 0; k < 5; k++) sum += result[i][j][k];
                    if (sum == 0) continue;
                    resultGrid.Rows.Add();
                    resultGrid.Rows[count].Cells[0].Value = lSchools[i];
                    resultGrid.Rows[count].Cells[1].Value = lNameSchools[i];
                    resultGrid.Rows[count].Cells[2].Value = 1 + gr_cnt;
                    resultGrid.Rows[count].Cells[3].Value = lNameSchools[j];
                    for (int k = 0; k < 5; k++)
                    {
                        /*if (statImpType == true) {
                            Result[lSchools[i]].Add(new Tuple<string, int, int, int>(lSchools[j], gr_cnt, k, result[i][j][k]));
                        }*/
                        resultGrid.Rows[count].Cells[4 + k].Value = result[i][j][k];
                    }

                    gr_cnt++;
                    count++;
                }
            }
            count = 0;
            resultGrid2.Rows.Clear();
            for (int j = 0; j < lSchools.Count; j++)
            { 
                int gr_cnt = 0;
                for (int i = 0; i < lSchools.Count; i++)
                {
                    if (i == j) continue;
                    int sum = 0;
                    for (int k = 0; k < 5; k++) sum += result[i][j][k];
                    if (sum == 0) continue;
                    resultGrid2.Rows.Add();
                    resultGrid2.Rows[count].Cells[0].Value = lSchools[j];
                    resultGrid2.Rows[count].Cells[1].Value = lNameSchools[j];
                    resultGrid2.Rows[count].Cells[2].Value = 1 + gr_cnt;
                    resultGrid2.Rows[count].Cells[3].Value = lNameSchools[i];
                    for (int k = 0; k < 5; k++)
                    {
                        /*if (statImpType == true)
                        {
                            Result2[lSchools[j]].Add(new Tuple<string, int, int, int>(lSchools[i], gr_cnt, k, result[i][j][k]));
                        }*/
                        resultGrid2.Rows[count].Cells[4 + k].Value = result[i][j][k];
                    }

                    gr_cnt++;
                    count++;
                }
            }
        }

        private void resultExpBtn_Click(object sender, EventArgs e)
        {
            expExcel(resultGrid, "DS phân công theo đơn vị");
        }

        private void delGrid_Click(object sender, EventArgs e)
        {
            statGrid.Rows.Clear();
        }

        private void delNeedGrid_Click(object sender, EventArgs e)
        {
            needGrid.Rows.Clear();
        }

        private void impStatBtn2_Click(object sender, EventArgs e)
        {
            statImpPath = showExcelDialog("DS đề xuất theo đơn vị");
            statImpType = true;
            statImpName.Text = "Chưa có file nào được chọn";
            statImpName2.Text = (statImpPath == "" ? "Chưa có file nào được chọn" : compressPath(statImpPath));
        }

        private void resultExpBtn2_Click(object sender, EventArgs e)
        {
            expExcel(resultGrid2, "DS phân công theo điểm thi");
        }

        private void allExpBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog file = new SaveFileDialog();
            file.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            file.Title = "DS phân công coi thi";
            file.ShowDialog();
            string allExpPath = file.FileName.ToString();
            if (allExpPath == "")
            {
                return;
            }
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);                
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            app.Visible = true;                


            void addSheet(DataGridView grid, string path, string sheetName, string title)
            {
                workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                worksheet = workbook.Sheets[sheetName];
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
            }            
            
            addSheet(resultGrid2, allExpPath, "Sheet1", "DS phân công theo điểm thi");
            addSheet(resultGrid, allExpPath, "Sheet1", "DS phân công theo đơn vị");
            addSheet(needGrid, allExpPath, "Sheet1", "DS thống kê theo điểm thi");
            addSheet(statFullGrid, allExpPath, "Sheet1", "DS đề xuất đầy đủ");
            addSheet(statGrid, allExpPath, "Sheet1", "DS đề xuất theo đơn vị"); 

            workbook.SaveAs(allExpPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            statGrid.Rows.Clear();
            statFullGrid.Rows.Clear();
        }

        private void statFullDelBtn_Click(object sender, EventArgs e)
        {
            statFullGrid.Rows.Clear();
        }

        private void statFullExpBtn_Click(object sender, EventArgs e)
        {
            expExcel(statFullGrid, "DS đề xuất đầy đủ");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 fr2 = new Form2();
            fr2.Show();
            fr2.arrangeGrid.Rows.Clear();
            if (!statImpType)
            {
                return;
            }
            int count = 0;
            foreach (KeyValuePair<string, List<string>[]> name in statNames)
            {
                int g = schools[name.Key];
                for (int j = 0; j < 5; j++) {
                    /*var rnd = new Random();
                    int n = name.Value[j].Count;
                    while (n > 1)
                    {
                        n--;
                        int k = rnd.Next(n + 1);
                        string value = name.Value[j][k];
                        name.Value[j][k] = name.Value[j][n];
                        name.Value[j][n] = value;
                    }*/
                    name.Value[j].Shuffle();
                }
                int gr_cnt = 0;
                for (int j = 0; j < 5; j++) 
                {
                    int cnt = 0;
                    foreach (KeyValuePair<string, int> school in schools) {
                        int l = school.Value;
                        for (int k = 0; k < result[g][l][j]; k++)
                        {
                            fr2.arrangeGrid.Rows.Add();
                            fr2.arrangeGrid.Rows[count].Cells[0].Value = name.Key;
                            fr2.arrangeGrid.Rows[count].Cells[1].Value = schoolNames[name.Key];
                            fr2.arrangeGrid.Rows[count].Cells[2].Value = name.Value[j][cnt];
                            fr2.arrangeGrid.Rows[count].Cells[3].Value = 1 + gr_cnt;
                            fr2.arrangeGrid.Rows[count].Cells[4].Value = schoolNames[school.Key];
                            fr2.arrangeGrid.Rows[count].Cells[5].Value = JobStr(j);
                            count++;
                            cnt++;
                        }
                        if (result[g][l][j] > 0)
                        {
                            gr_cnt++;
                        }
                    }
                }
            }
        }

        private void statAllExpBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog file = new SaveFileDialog();
            file.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            file.Title = "DS đề xuất theo đơn vị";
            file.ShowDialog();
            string allExpPath = file.FileName.ToString();
            if (allExpPath == "")
            {
                return;
            }
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            app.Visible = true;


            void addSheet(DataGridView grid, string path, string sheetName, string title)
            {
                workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                worksheet = workbook.Sheets[sheetName];
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
            }

            addSheet(statFullGrid, allExpPath, "Sheet1", "DS đề xuất đầy đủ");
            addSheet(statGrid, allExpPath, "Sheet1", "DS đề xuất theo đơn vị");

            workbook.SaveAs(allExpPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.Quit();
        }

        void copySampleFile(string title, string fileName)
        {
            SaveFileDialog file = new SaveFileDialog();
            file.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            file.FileName = fileName;
            file.Title = title;
            file.ShowDialog();
            string samplePath = file.FileName.ToString();
            if (samplePath == fileName)
            {
                return;
            }
            File.Copy(Application.StartupPath + "\\samples\\" + fileName, samplePath, true);
        }

        private void statSampleBtn_Click(object sender, EventArgs e)
        {
            /*FolderBrowserDialog folder = new FolderBrowserDialog();
            //file.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            //folder.Description = "DS đề xuất mẫu";
            folder.ShowDialog();
            string samplePath = objToString(folder.SelectedPath);*/
            copySampleFile("DS đề xuất theo đơn vị", "1_DanhSachDeXuatCoiThi.xlsx");
        }

        private void statSampleBtn2_Click(object sender, EventArgs e)
        {
            copySampleFile("DS thống kê theo đơn vị", "2_DanhSachThongKeDeXuat.xlsx");
        }

        private void needSampleBtn_Click(object sender, EventArgs e)
        {
            copySampleFile("DS thống kê theo điểm thi", "3_DanhSachThongKeDiemThi.xlsx");
        }

        private void resultExpBtn_Click_1(object sender, EventArgs e)
        {
            expExcel(statGrid, "DS phân công theo đơn vị");
        }

        private void resultExpBtn2_Click_1(object sender, EventArgs e)
        {

            expExcel(statGrid, "DS phân công theo điểm thi");
        }
    }

    static class ExtensionsClass
    {
        private static Random rng = new Random();

        public static void Shuffle<T>(this IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
    }

}