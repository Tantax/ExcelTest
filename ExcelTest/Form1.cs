using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelClassLibrary;

namespace ExcelTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            String ConnectString = @"Data Source=.;Initial Catalog=ExcelTestDB;Persist Security Info=True;User ID=sa;Password=123456";
            this.textBoxExcelPath.AllowDrop = true;//允许拖拽
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO:  这行代码将数据加载到表“excelTestDBDataSet1.UserInformation”中。您可以根据需要移动或删除它。
            this.userInformationTableAdapter1.Fill(this.excelTestDBDataSet1.UserInformation);

        }

        #region TextBox的拖拽

        private void textBoxExcelPath_DragDrop(object sender, DragEventArgs e)
        {
            this.textBoxExcelPath.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
        }

        private void textBoxExcelPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        #endregion

        private void buttonInExcel_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(this.textBoxExcelPath.Text))
            {
                MessageBox.Show("请拖入文件到Text框来读取地址", "提示", MessageBoxButtons.OK);
            }
            else
            {
                String excelUrl = this.textBoxExcelPath.Text;
                DataTable dataTable = GetDgvToTable(this.dataGridView1);
                //DataTable dataTable = new DataTable("UserInformation");
                
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(dataTable);
                var list = ExcelHelper.DataSetToList<ExcelModel>(dataSet, 0);

                try
                {
                    //ExcelHandler.DSToExcel2003(excelUrl, dataTable.DataSet, "UserInformation");
                    ExcelHelper.SaveToExcel<ExcelModel>(list, "ExcelTest");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("错误信息：" + ex.Message, "警告", MessageBoxButtons.OK);
                }
            }
        }

        /// <summary>
        /// 读取DataGridView中的数据
        /// </summary>
        /// <param name="dgv">DataGridView对象</param>
        /// <returns>DataTable对象</returns>
        private DataTable GetDgvToTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        
    }
}
