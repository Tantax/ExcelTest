using System;
using System.Xml;

namespace Common
{
    public class SellException : Exception
    {
        private DateTime dt;
        /// <summary>
        /// 发生异常时间
        /// </summary>
        public DateTime Dt
        {
            get { return dt; }
            set { dt = value; }
        }
        private string errCode;
        /// <summary>
        /// 发生异常代码
        /// </summary>
        public string ErrCode
        {
            get { return errCode; }
            set { errCode = value; }
        }
        private string winName;
        /// <summary>
        /// 发生异常窗体
        /// </summary>
        public string WinName
        {
            get { return winName; }
            set { winName = value; }
        }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="_errCode"></param>
        /// <param name="_winName"></param>
        /// <param name="ex"></param>
        public SellException(string _errCode, string _winName, Exception ex)
            : base(ex.Message)
        {
            this.dt = DateTime.Now;
            this.errCode = _errCode;
            this.winName = _winName;
        }
        /// <summary>
        /// 把异常的信息写入XML
        /// </summary>
        public void WriteXml()
        {
            //相对路径
            string saveFilePath = "XMLException.xml";
            //实例化XML文档对象
            XmlDocument doc = new XmlDocument();
            //把文档载入到内存
            doc.Load(saveFilePath);
            //读取根元素
            XmlElement root = doc.DocumentElement;
            //设置属性值
            var d = doc.CreateAttribute("id");
            if (!root.HasChildNodes)//如果根元素下没有任何节点
            {
                d.Value = "0";
            }
            else//有节点
            {
                //取出根元素最后一个子元素的id属性值，然后对其++
                XmlElement lastErrInfo = root.LastChild as XmlElement;
                int lastId = Convert.ToInt32(lastErrInfo.GetAttribute("id"));
                lastId++;
                d.Value = lastId.ToString();
            }
            //创建root下的子元素
            XmlElement ErrInfo = doc.CreateElement("ErrInfo");
            //添加属性
            ErrInfo.Attributes.Append(d);
            //创建ErrInfo下的子元素
            XmlElement errDate = doc.CreateElement("errDate");
            errDate.InnerText = this.dt.ToString();
            XmlElement errFile = doc.CreateElement("errFile");
            errFile.InnerText = this.winName;
            XmlElement errCode = doc.CreateElement("errCode");
            errCode.InnerText = this.errCode;
            XmlElement errMsg = doc.CreateElement("errMsg");
            errMsg.InnerText = this.Message;
            //把各个元素添加到ErrInfo下
            ErrInfo.AppendChild(errDate);
            ErrInfo.AppendChild(errFile);
            ErrInfo.AppendChild(errCode);
            ErrInfo.AppendChild(errMsg);
            //把ErrInfo添加到root下
            root.AppendChild(ErrInfo);
            //把内存中的doc存入到文档中
            doc.Save(saveFilePath);
        }
        /// <summary>
        /// 显示异常详细信息
        /// </summary>
        /// <returns></returns>
        public string ShowErrMsg()
        {
            return "异常发生的时间：" + dt + "\t\n异常发生的窗体：" + winName + "\t\n异常代码：" + errCode + "\t\n异常信息：" + Message;
        }

    }
}
