using Dal;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTest
{
    public partial class Form1 : Form
    {
        SQLHelper sqlHelper = null;
        public Form1()
        {
            InitializeComponent();
            if (sqlHelper == null)
            {
                sqlHelper = new SQLHelper();
            }
            this.textBoxExcelPath.AllowDrop = true;//允许拖拽
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 这行代码将数据加载到表“excelTestDBDataSet1.UserInformation”中。您可以根据需要移动或删除它。
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

        /// <summary>
        /// 数据库导出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonInExcel_Click(object sender, EventArgs e)
        {
            if (IsInput(this.textBoxExcelPath))
            {
                String excelUrl = this.textBoxExcelPath.Text;
                //DataTable dataTable = GetDgvToTable(this.dataGridView1);

                //DataSet dataSet = new DataSet();
                //dataSet.Tables.Add(dataTable);

                DataSet dataSet = sqlHelper.GetDataSet(@"SELECT 
	                                                Id,
	                                                PaymentAccount,
	                                                CustomerName,
	                                                Adress,
	                                                PhoneNumber,
	                                                GasType,
	                                                FactoryNumber,
	                                                Area,
	                                                Community,
	                                                Floor
                                                FROM UserInformation");
                var list = ExcelHelper.DataSetToList<ExcelModel>(dataSet, 0);

                try
                {
                    ExcelHelper.SaveToExcel<ExcelModel>(list, excelUrl);
                    MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("错误信息：" + ex.Message, "警告", MessageBoxButtons.OK);
                }
            }
        }

        /// <summary>
        /// Excel导入数据库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonOutExcel_Click(object sender, EventArgs e)
        {
            if (IsInput(this.textBoxExcelPath))
            {
                DateTime transactionStart, transactionEnd, loadFromExcelStart, loadFrmExcelEnd;
                loadFromExcelStart = DateTime.Now;
                var excelList = ExcelHelper.LoadFromExcel<ExcelModel>(this.textBoxExcelPath.Text);
                loadFrmExcelEnd = DateTime.Now;
                this.richTextBoxLog.AppendText("LoadFromExcel函数执行时间为：" + (loadFrmExcelEnd - loadFromExcelStart).ToString() + Environment.NewLine);
                //List<SqlParameter> spList = new List<SqlParameter>();
                SqlParameter sp0 = new SqlParameter("@paymentaccount", SqlDbType.BigInt);
                SqlParameter sp1 = new SqlParameter("@customername", SqlDbType.NVarChar, 10);
                SqlParameter sp2 = new SqlParameter("@adress", SqlDbType.NVarChar);
                SqlParameter sp3 = new SqlParameter("@phonenumber", SqlDbType.VarChar, 30);
                SqlParameter sp4 = new SqlParameter("@gastype", SqlDbType.NVarChar, 50);
                SqlParameter sp5 = new SqlParameter("@factorynumber", SqlDbType.NChar, 20);
                SqlParameter sp6 = new SqlParameter("@area", SqlDbType.NVarChar, 50);
                SqlParameter sp7 = new SqlParameter("@community", SqlDbType.NVarChar, 50);
                SqlParameter sp8 = new SqlParameter("@floor", SqlDbType.NVarChar, 50);
                // 获取数据库连接
                SqlConnection connection = sqlHelper.GetConnectionObject();
                SqlCommand command = connection.CreateCommand();
                SqlTransaction transaction;
                transactionStart = DateTime.Now;
                //启动事务
                transaction = connection.BeginTransaction("SampleTransaction");

                //设定SqlCommand的事务和连接对象
                //command.Connection = connection;
                command.Transaction = transaction;
                try
                {
                    int cntLog = 0;// 插入数据计数
                    string sql = @"INSERT INTO UserInformation
	                                (
	                                PaymentAccount,
	                                CustomerName,
	                                Adress,
	                                PhoneNumber,
	                                GasType,
	                                FactoryNumber,
	                                Area,
	                                Community,
	                                Floor
	                                )
                                VALUES 
	                                (
	                                @paymentaccount,
	                                @customername,
	                                @adress,
	                                @phonenumber,
	                                @gastype,
	                                @factorynumber,
	                                @area,
	                                @community,
	                                @floor
	                                )";
                    command.CommandText = sql;
                    SqlParameter[] sqlParameterArray = new SqlParameter[] { sp0, sp1, sp2, sp3, sp4, sp5, sp6, sp7, sp8 };
                    foreach (var item in excelList)
                    {
                        sqlParameterArray[0].Value = item.PaymentAccount;
                        sqlParameterArray[1].Value = item.CustomerName;
                        sqlParameterArray[2].Value = item.Adress;
                        sqlParameterArray[3].Value = item.PhoneNumber == null ? "" : item.PhoneNumber;
                        sqlParameterArray[4].Value = item.GasType == null ? "" : item.GasType;
                        sqlParameterArray[5].Value = item.FactoryNumber == null ? "" : item.FactoryNumber;
                        sqlParameterArray[6].Value = item.Area;
                        sqlParameterArray[7].Value = item.Community;
                        sqlParameterArray[8].Value = item.Floor;
                        //spList.Clear();  // 清空List
                        //spList.Add(sp0);
                        //spList.Add(sp1);
                        //spList.Add(sp2);
                        //spList.Add(sp3);
                        //spList.Add(sp4);
                        //spList.Add(sp5);
                        //spList.Add(sp6);
                        //spList.Add(sp7);
                        //spList.Add(sp8);
                        //SqlCommand sqlCmd = sqlHelper.GetCommand(sql, spList);
                        //int result = sqlCmd.ExecuteNonQuery();
                        //sqlHelper.GetExecuteNonQuery(sql, spList);
                        command.Parameters.Clear();
                        command.Parameters.AddRange(sqlParameterArray);
                        cntLog += command.ExecuteNonQuery();

                    }
                    // 完成提交
                    transaction.Commit();
                    transactionEnd = DateTime.Now;
                    this.richTextBoxLog.AppendText("完成提交。" + Environment.NewLine);
                    this.richTextBoxLog.AppendText("共插入" + cntLog + "条数据" + Environment.NewLine);
                    this.richTextBoxLog.AppendText("总共耗时：" + (transactionEnd - transactionStart).ToString() + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    this.richTextBoxLog.AppendText("出错，错误信息：" + ex.Message + Environment.NewLine + "数据回滚..." + Environment.NewLine);
                    //数据回滚
                    transaction.Rollback();
                }
                finally
                {
                    sqlHelper.CloseConnection();
                    this.richTextBoxLog.AppendText("结束。" + Environment.NewLine);
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
            for (int row = 0; row < dgv.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = dgv.Rows[row].Cells[countsub].Value;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private bool IsInput(TextBox textBox)
        {
            if (String.IsNullOrEmpty(textBox.Text))
            {
                MessageBox.Show("请拖入文件到Text框来读取地址", "提示", MessageBoxButtons.OK);
                return false;
            }
            else
            {
                return true;
            }
        }

    }
}
