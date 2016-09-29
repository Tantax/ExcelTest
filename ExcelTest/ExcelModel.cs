using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest
{
   public class ExcelModel
    {
       [DisplayName("ID")]
       public int Id { get; set; }
       [DisplayName("账户")]
       public string PaymentAccount { get; set; }
       [DisplayName("用户名")]
       public string CustomerName { get; set; }
       [DisplayName("地址")]
       public string Adress { get; set; }
       [DisplayName("电话")]
       public string PhoneNumber { get; set; }
       [DisplayName("类型")]
       public string GasType { get; set; }
       [DisplayName("公司电话")]
       public string FactoryNumber { get; set; }
       [DisplayName("区域")]
       public string Area { get; set; }
       [DisplayName("小区")]
       public string Community { get; set; }
       [DisplayName("楼栋")]
       public string Floor { get; set; }


    }
}
