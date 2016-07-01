using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Common;

namespace Dal
{
    public class SQLHelper : AbstractSQLHelper
    {
        /// <summary>
        /// 返回第一行第一列的值
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="pars"></param>
        /// <returns></returns>
        public override object GetSingleData(string sql, params SqlParameter[] pars)
        {
            SqlCommand cmd = GetCommand(sql, pars);
            object o = cmd.ExecuteScalar();
            CloseConnection();
            return o;
        }
        /// <summary>
        /// 增删改查
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="pars"></param>
        /// <returns></returns>
        public override bool GetExecuteNonQuery(string sql, params SqlParameter[] pars)
        {
            try
            {
                SqlCommand cmd = GetCommand(sql, pars);
                int result = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                CloseConnection();
                return result > 0 ? true : false;
            }
            catch (Exception ex)
            {

                throw new SellException("SQLHelper36", "", ex);
            }
        }
        public override bool GetExecuteNonQuery(string sql, List<SqlParameter> pars)
        {
            try
            {
                SqlCommand cmd = GetCommand(sql, pars);
                int result = cmd.ExecuteNonQuery();
                CloseConnection();
                return result > 0 ? true : false;
            }
            catch (Exception ex)
            {

                throw new SellException("SQLHelper59", "", ex);
            }
        }
        public override bool GetExecuteNonQuery(string sql)
        {
            try
            {
                SqlCommand cmd = GetCommand(sql);
                int result = cmd.ExecuteNonQuery();
                CloseConnection();
                return result > 0 ? true : false;
            }
            catch (Exception ex)
            {

                throw new SellException("SQLHelper67", "", ex);
            }
        }
        /// <summary>
        /// 返回一个SqlDataReader（使用完成之后要用sqlHelper.CloseConnection()关闭数据连接）
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="pars"></param>
        /// <returns></returns>
        public override SqlDataReader GetDataReader(string sql, params SqlParameter[] pars)
        {
            SqlCommand cmd = GetCommand(sql, pars);
            SqlDataReader sdr = cmd.ExecuteReader();
            return sdr;
        }
        public override SqlDataReader GetDataReader(string sql)
        {
            SqlCommand cmd = GetCommand(sql);
            SqlDataReader sdr = cmd.ExecuteReader();
            return sdr;
        }
        /// <summary>
        /// 返回一个GetDataSet
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="pars"></param>
        /// <returns></returns>
        public override DataSet GetDataSet(string sql, List<SqlParameter> pars)
        {
            SqlCommand cmd = GetCommand(sql, pars);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            CloseConnection();
            return ds;
        }
        public override DataSet GetDataSet(string sql, params SqlParameter[] pars)
        {
            SqlCommand cmd = GetCommand(sql, pars);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            CloseConnection();
            return ds;
        }
        public override DataSet GetDataSet(string sql)
        {
            SqlCommand cmd = GetCommand(sql);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            CloseConnection();
            return ds;
        }

       
    }
}