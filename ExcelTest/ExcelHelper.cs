using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest
{
    public class ExcelHelper
    {


        #region 保存数据列表到Excel（泛型）
        /// <summary>
        /// 保存数据列表到Excel（泛型）
        /// </summary>
        /// <typeparam name="T">集合数据类型</typeparam>
        /// <param name="data">数据列表</param>
        /// <param name="FilePath">Excel文件路径</param>
        /// <param name="OpenPassword">Excel打开密码</param>
        public static void SaveToExcel<T>(IEnumerable<T> data, string FilePath, string OpenPassword = "")
        {
            FileInfo file = new FileInfo(FilePath);
            if (file.Exists)
            {
                file.Delete();//保证只有一个workbook，不然后面会报错
                file = new FileInfo(FilePath);
            }            
            try
            {
                using (ExcelPackage ep = new ExcelPackage(file))
                {
                    ExcelWorksheet ws = ep.Workbook.Worksheets.Add(typeof(T).Name);
                    ws.Cells["A1"].LoadFromCollection(data, true, TableStyles.Medium10);

                    ep.Save();
                }
            }
            catch (InvalidOperationException ex)
            {
                throw ex;
            }
        }
        #endregion




        #region 从Excel中加载数据（泛型）
        /// <summary>
        /// 从Excel中加载数据（泛型）
        /// </summary>
        /// <typeparam name="T">每行数据的类型</typeparam>
        /// <param name="FilePath">Excel文件路径</param>
        /// <returns>泛型列表</returns>
        public static IEnumerable<T> LoadFromExcel<T>(string FilePath) where T : new()
        {
            FileInfo existingFile = new FileInfo(FilePath);
            List<T> resultList = new List<T>();
            Dictionary<string, int> dictHeader = new Dictionary<string, int>();

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                int colStart = worksheet.Dimension.Start.Column;  //工作区开始列
                int colEnd = worksheet.Dimension.End.Column;       //工作区结束列
                int rowStart = worksheet.Dimension.Start.Row;       //工作区开始行号
                int rowEnd = worksheet.Dimension.End.Row;       //工作区结束行号

                //将每列标题添加到字典中
                for (int i = colStart; i <= colEnd; i++)
                {
                    dictHeader[worksheet.Cells[rowStart, i].Value.ToString()] = i;
                }

                List<PropertyInfo> propertyInfoList = new List<PropertyInfo>(typeof(T).GetProperties());

                for (int row = rowStart + 1; row <= rowEnd; row++)
                {
                    T result = new T();

                    //为对象T的各属性赋值
                    foreach (PropertyInfo p in propertyInfoList)
                    {
                        try
                        {
                            ExcelRange cell = worksheet.Cells[row, dictHeader[p.Name]]; //与属性名对应的单元格

                            if (cell.Value == null)
                                continue;
                            switch (p.PropertyType.Name.ToLower())
                            {
                                case "string":
                                    p.SetValue(result, cell.GetValue<String>());
                                    break;
                                case "int16":
                                    p.SetValue(result, cell.GetValue<Int16>());
                                    break;
                                case "int32":
                                    p.SetValue(result, cell.GetValue<Int32>());
                                    break;
                                case "int64":
                                    p.SetValue(result, cell.GetValue<Int64>());
                                    break;
                                case "decimal":
                                    p.SetValue(result, cell.GetValue<Decimal>());
                                    break;
                                case "double":
                                    p.SetValue(result, cell.GetValue<Double>());
                                    break;
                                case "datetime":
                                    p.SetValue(result, cell.GetValue<DateTime>());
                                    break;
                                case "boolean":
                                    p.SetValue(result, cell.GetValue<Boolean>());
                                    break;
                                case "byte":
                                    p.SetValue(result, cell.GetValue<Byte>());
                                    break;
                                case "char":
                                    p.SetValue(result, cell.GetValue<Char>());
                                    break;
                                case "single":
                                    p.SetValue(result, cell.GetValue<Single>());
                                    break;
                                default:
                                    break;
                            }
                        }
                        catch (KeyNotFoundException ex)
                        { }
                    }
                    resultList.Add(result);
                }
            }
            return resultList;
        }
        #endregion

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
            // 在这里写 获取T类型的所有公有属性。 注意这里仅仅是获取T类型的公有属性，不是公有方法，也不是公有字段，当然也不是私有属性                                                 
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
                                // TODO:模型类型和DataTable的类型不一致，如一个是int，一个是string
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
