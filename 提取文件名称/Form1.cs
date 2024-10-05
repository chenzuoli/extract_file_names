using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace 提取文件名称
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnClickThis(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowse = new FolderBrowserDialog();
            if (folderBrowse.ShowDialog() == DialogResult.OK)
            {
                String[] files = Directory.GetFiles(folderBrowse.SelectedPath);

                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("文件夹", typeof(string));
                dt.Columns.Add("文件名称", typeof(string));

                for (int i = 0; i < files.Length; i++)
                {
                    dt.Rows.Add(Path.GetDirectoryName(files[i]),  Path.GetFileName(files[i]));
                }

                // 将DataTable数据绑定到DataGridView
                dataGridView1.RowHeadersVisible = true;
                dataGridView1.DataSource = dt;
                // 让datatable自动填满整个DataGridView
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }

        private void btnConfirm(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("无法创建Excel，您可能需要安装Excel");
                return;
            }

            // 创建excel工作薄
            Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
            Worksheet workSheet = null;

            // 创建工作表
            workSheet = workBook.Sheets["Sheet1"];
            workSheet = workBook.ActiveSheet;

            // 表头
            Range headerRow = workSheet.Rows[1];
            headerRow.Cells[1, 1] = "文件路径";
            headerRow.Cells[1, 2] = "文件名";
            // 表头格式
            headerRow.Font.Bold = true;
            headerRow.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);


            // 将DataGridView表格内数据复制到excel工作表
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    workSheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }

            // 导出到excel文件
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 Workbook (*.xls)|*.xls";
            saveFileDialog.Title = "保存文件名称到Excel";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                try
                {
                    workBook.SaveAs(saveFileDialog.FileName);
                    workBook.Close(false);
                    excelApp.Quit();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                    workSheet = null;
                    workBook = null;
                    excelApp = null;

                    MessageBox.Show("Excel文件已保存到：" + saveFileDialog.FileName);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("保存文件名称到Excel失败，请稍后重试。" + ex.Message);
                }
            }
        }

        // 注册事件 this.dataGridView1.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridView1_CellPainting);


        private void dataGridView1_CellPainting(Object sender, DataGridViewCellPaintingEventArgs e)
        {

            if (e.RowIndex >= 0 && e.ColumnIndex == -1)
            {
                e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);
                using (Brush brush  = new SolidBrush(e.CellStyle.ForeColor))
                {
                    e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.CellStyle.Font, brush, e.CellBounds.Location.X + 14, e.CellBounds.Location.Y + 8);
                }
                e.Handled = true;
            }
        }

    }
}
