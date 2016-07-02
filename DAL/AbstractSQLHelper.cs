using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace Dal
{
    public abstract class AbstractSQLHelper
    {
        public static readonly string strCon = @"Data Source=.;Initial Catalog=ExcelTestDB;Persist Security Info=True;User ID=sa;Password=123456";
        private SqlConnection con = null;
        /// <summary>
        /// 打开连接
        /// </summary>
        private void OpenConnection()
        {
            if (con == null)
            {
                con = new SqlConnection(strCon);
            }            
            if (con.State != ConnectionState.Open)
            {
                con.Close();
                con.Open();
            }
        }
        public SqlConnection GetConnectionObject()
        {
            OpenConnection();
            return con;
        }
        /// <summary>
        /// 关闭连接
        /// </summary>
        public void CloseConnection()
        {
            if (con != null)
            {
                con.Close();
                con.Dispose();
            }
        }
        /// <summary>
        /// 返回Command对象
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="pars"></param>
        /// <returns></returns>
        public SqlCommand GetCommand(string sql, params SqlParameter[] pars)
        {
            OpenConnection();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = sql;
            if (pars != null)
            {
                cmd.Parameters.AddRange(pars);
            }
            return cmd;
        }
        public SqlCommand GetCommand(string sql, List<SqlParameter> pars)
        {
            OpenConnection();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = sql;
            if (pars != null)
            {
                foreach (SqlParameter s in pars)
                {
                    cmd.Parameters.Add(s);
                }
            }
            return cmd;
        }
        public SqlCommand GetCommand(string sql)
        {
            OpenConnection();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = sql;
            return cmd;
        }

        public abstract object GetSingleData(string sql, params SqlParameter[] pars);
        public abstract bool GetExecuteNonQuery(string sql, params SqlParameter[] pars);
        public abstract bool GetExecuteNonQuery(string sql, List<SqlParameter> pars);
        public abstract bool GetExecuteNonQuery(string sql);
        public abstract SqlDataReader GetDataReader(string sql, params SqlParameter[] pars);
        public abstract SqlDataReader GetDataReader(string sql);
        public abstract DataSet GetDataSet(string sql, List<SqlParameter> pars);
        public abstract DataSet GetDataSet(string sql, params SqlParameter[] pars);
        public abstract DataSet GetDataSet(string sql);

    }
}