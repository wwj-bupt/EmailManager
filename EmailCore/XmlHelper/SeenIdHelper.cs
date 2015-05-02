using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace EmailCore.XmlHelper
{
    /// <summary>
    /// 操作已读邮件列表的帮助类
    /// </summary>
    public class SeenIdHelper
    {
        private string fileName = "seenid.xml";
        private string fullFileName;
        public SeenIdHelper()
        {
            fullFileName = Environment.CurrentDirectory + "/" + fileName;
        }

        public void Save(List<string> ids)
        {
            var xml = new XElement("MessageIds");
            foreach (var id in ids)
            {
                xml.Add(new XElement("Id", id));
            }
            xml.Save(fullFileName);
        }

        public List<string> Load()
        {
            try
            {
                var xml = XElement.Load(fullFileName);
                var ids = new List<string>();
                foreach (XElement node in xml.Elements("Id"))
                {
                    ids.Add(node.Value);
                }
                return ids;
            }
            catch
            {
                return new List<string>();
            }
        }
    }
}
