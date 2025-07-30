using System;
using System.Data;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;
using System.Linq;

namespace ExcelEditorApp
{
    public partial class Form1 : Form
    {
        private string filePath;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx;*.xls";


            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheets.Worksheet(1);

                    DataTable dt = new DataTable();

                    var range = worksheet.RangeUsed();
                    bool firstRow = true;

                    foreach (var row in range.Rows())
                    {
                        if (firstRow)
                        {
                            foreach (var cell in row.Cells())
                                dt.Columns.Add(cell.GetString());
                            firstRow = false;
                        }
                        else
                        {
                            dt.Rows.Add();
                            for (int i = 0; i < row.Cells().Count(); i++)
                                dt.Rows[dt.Rows.Count - 1][i] = row.Cell(i + 1).Value;
                        }
                    }

                    dataGridView1.DataSource = dt;
                }
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            Random rand = new Random();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                if (row.Cells.Count >= 7 && row.Cells[4].Value != null)
                {
                    object value = row.Cells[4].Value;

                    if (double.TryParse(value.ToString(), out double number))
                    {
                        bool putInSixth = rand.Next(0, 2) == 0;

                        if (putInSixth)
                        {
                            row.Cells[5].Value = number; 
                            row.Cells[6].Value = 0;      
                        }
                        else
                        {
                            row.Cells[5].Value = 0;
                            row.Cells[6].Value = number;
                        }
                    }
                }
            }
            dataGridView1.EndEdit();
            dataGridView1.RefreshEdit();

            MessageBox.Show("تم التعديل بنجاح ✅", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("لا يوجد بيانات للتصدير");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Workbook|*.xlsx";
            sfd.Title = "احفظ الملف كـ Excel";
            sfd.FileName = "تقرير.xlsx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (XLWorkbook workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("البيانات");

                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        ws.Cell(1, i + 1).Value = dataGridView1.Columns[i].HeaderText;
                        ws.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                        ws.Cell(1, i + 1).Style.Font.Bold = true;
                        ws.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].IsNewRow) continue;

                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            var cellValue = dataGridView1.Rows[i].Cells[j].Value;
                            ws.Cell(i + 2, j + 1).Value = cellValue != null ? cellValue.ToString() : "";
                        }
                    }


                    ws.Columns().AdjustToContents();

                    workbook.SaveAs(sfd.FileName);
                    MessageBox.Show("تم حفظ الملف بنجاح!");
                }
            }
        }


    }
}
