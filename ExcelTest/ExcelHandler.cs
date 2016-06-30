using System;
using System.Data;
using System.Data.OleDb;

namespace ExcelClassLibrary
{
    class ExcelHandler
    {
        /// <summary>
        /// 导出数据到Excel
        /// </summary>
        /// <param name="Path">需要导入的Excel地址</param>
        /// <param name="oldds">需要导入的数据</param>
        /// <param name="TableName">表名</param>
        public static void DSToExcel2003(string Path, DataSet oldds, string TableName)
        {
            //Excel2003的连接字符串
            string strCon = @"Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source =" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
            //执行导入
            ExcuteSQL(oldds, TableName, strCon);


        }

        /// <summary>
        /// 导出数据到Excel
        /// </summary>
        /// <param name="Path">需要导入的Excel地址</param>
        /// <param name="oldds">需要导入的数据</param>
        /// <param name="TableName">表名</param>
        public static void DSToExcel2007(string Path, DataSet oldds, string TableName)
        {
            //Excel2007的连接字符串  
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + Path + ";" + "Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
            //执行导入
            ExcuteSQL(oldds, TableName, strConn);


        }
        /// <summary>
        /// 执行
        /// </summary>
        /// <param name="oldds">需要导入的数据</param>
        /// <param name="TableName">表名</param>
        /// <param name="strCon">连接字符串</param>
        private static void ExcuteSQL(DataSet oldds, string TableName, string strCon)
        {
            //连接
            OleDbConnection myConn = new OleDbConnection(strCon);

            string strCom = "select * from [" + TableName + "$]";

            try
            {
                myConn.Open();
                OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
                System.Data.OleDb.OleDbCommandBuilder builder = new OleDbCommandBuilder(myCommand);

                //QuotePrefix和QuoteSuffix主要是对builder生成InsertComment命令时使用。   
                //获取insert语句中保留字符（起始位置）  
                builder.QuotePrefix = "[";

                //获取insert语句中保留字符（结束位置）   
                builder.QuoteSuffix = "]";

                DataSet newds = new DataSet();
                //获得表结构
                DataTable ndt = oldds.Tables[0].Clone();
                //清空数据
                //ndt.Rows.Clear();

                ndt.TableName = TableName;
                newds.Tables.Add(ndt);

                //myCommand.Fill(newds, TableName);

                for (int i = 0; i < oldds.Tables[0].Rows.Count; i++)
                {
                    //在这里不能使用ImportRow方法将一行导入到news中，
                    //因为ImportRow将保留原来DataRow的所有设置(DataRowState状态不变)。
                    //在使用ImportRow后newds内有值，但不能更新到Excel中因为所有导入行的DataRowState!=Added     
                    DataRow nrow = newds.Tables[0].NewRow();
                    for (int j = 0; j < oldds.Tables[0].Columns.Count; j++)
                    {
                        nrow[j] = oldds.Tables[0].Rows[i][j];
                    }
                    newds.Tables[0].Rows.Add(nrow);
                }
                //插入数据
                myCommand.Update(newds, TableName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                myConn.Close();
            }
        }

        private DataSet importExcelToDataSet(string FilePath)
        {
            string strConn;
            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + FilePath + ";Extended Properties=Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            OleDbDataAdapter myCommand = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", strConn);
            DataSet myDataSet = new DataSet();
            try
            {
                myCommand.Fill(myDataSet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return myDataSet;
        }
        

		
    }
}