using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dal;
using System.Data;
using System.Reflection;

namespace AddressHandler
{
    class Program
    {
        static void Main(string[] args)
        {
        }
        static void AddressHandler()
        {
            
        }

        static List<AddressModel> GetAreaName()
        {
            string sql = @"SELECT distinct [Area]
  FROM [ExcelTestDB].[dbo].[UserInformation]";
            SQLHelper sqlHelper = new SQLHelper();
            DataSet ds = sqlHelper.GetDataSet(sql);
            return DataSetToList<AddressModel>(ds, 0);
            return null;
        }
        static List<AddressModel> GetCommunity()
        {
            return null;
        }
        static List<AddressModel> GetFloor()
        {
            return null;
        }

        #region DataSet转List
        /// <summary>  
        /// DataSetToList  
        /// </summary>  
        /// <typeparam name="T">转换类型</typeparam>  
        /// <param name="ds">一个DataSet实例，也就是数据源</param>  
        /// <param name="tableIndext">DataSet容器里table的下标，只有用于取得哪个table，也就是需要转换表的索引</param>  
        /// <returns></returns>  
        public static List<T> DataSetToList<T>(DataSet ds, int tableIndext)
        {
            //确认参数有效  
            if (ds == null || ds.Tables.Count <= 0 || tableIndext < 0)
            {
                return null;
            }
            DataTable dt = ds.Tables[tableIndext]; //取得DataSet里的一个下标为tableIndext的表，然后赋给dt  

            IList<T> list = new List<T>();  //实例化一个list  
            // 获取T类型的所有公有属性。 注意这里仅仅是获取T类型的公有属性，不是公有方法，也不是公有字段，当然也不是私有属性                                                 
            PropertyInfo[] tMembersAll = typeof(T).GetProperties();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //创建泛型对象  
                T t = Activator.CreateInstance<T>();


                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //遍历tMembersAll  
                    foreach (PropertyInfo tMember in tMembersAll)
                    {
                        //取dt表中j列的名字，并把名字转换成大写的字母。整条代码的意思是：如果列名和属性名称相同时赋值  
                        if (dt.Columns[j].ColumnName.ToUpper().Equals(tMember.Name.ToUpper()))
                        {
                            //dt.Rows[i][j]表示取dt表里的第i行的第j列；DBNull是指数据库中当一个字段没有被设置值的时候的值，相当于数据库中的“空值”。   
                            if (dt.Rows[i][j] != DBNull.Value)
                            {
                                //SetValue是指：将指定属性设置为指定值。 tMember是T泛型对象t的一个公有成员，整条代码的意思就是：将dt.Rows[i][j]赋值给t对象的tMember成员,参数详情请参照http://msdn.microsoft.com/zh-cn/library/3z2t396t(v=vs.100).aspx/html
                                tMember.SetValue(t, dt.Rows[i][j], null);

                            }
                            else
                            {
                                tMember.SetValue(t, null, null);
                            }
                            break;//注意这里的break是写在if语句里面的，意思就是说如果列名和属性名称相同并且已经赋值了，那么我就跳出foreach循环，进行j+1的下次循环  
                        }
                    }
                }

                list.Add(t);
            }
            return list.ToList();

        }
        #endregion
    }
}
