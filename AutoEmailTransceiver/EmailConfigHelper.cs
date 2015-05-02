using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace AutoEmailTransceiver
{
    public class EmailConfigHelper
    {
        private static List<EmailConfig> EmailConfigs;
        /// <summary>
        /// 回复的标题
        /// </summary>
        public static string Subject { get; private set; }
        /// <summary>
        /// 获取邮件服务器中邮件的时间间隔
        /// </summary>
        public static int ServerInterval { get; private set; }
        static EmailConfigHelper()
        {
            var xml = LoadConfig();
            InitialEmailConfigs(xml);
        }
        /// <summary>
        /// 获取选择的邮件配置信息
        /// </summary>
        /// <param name="name">要获取的邮件配置名称，如QQ,163,Sina等</param>
        /// <returns></returns>
        public static EmailConfig GetEmailConfig(string name)
        {
            return EmailConfigs.FirstOrDefault(it => it.Name == name);
        }

        /// <summary>
        /// 获取邮件配置名称
        /// </summary>
        public static IEnumerable<string> EmailNames
        {
            get
            {
                return EmailConfigs.Select(it => it.Name);
            }
        }

        private static XElement LoadConfig()
        {
            var configName = Environment.CurrentDirectory + "/emailconfig.xml";
            return XElement.Load(configName);
        }

        private static void InitialEmailConfigs(XElement xml)
        {
            EmailConfigs = new List<EmailConfig>();
            Subject = xml.Element("Subject").Value;
            int interval = 1;
            //获取定时器的时间间隔
            ServerInterval = int.TryParse(xml.Element("ServerInterval").Value, out interval) ? interval : 1;
            foreach (XElement child in xml.Elements("Email"))
            {
                var config = new EmailConfig()
                {
                    Name = child.Attribute("Name").Value,
                    Pop3Server = child.Element("Pop3Server").Value,
                    Pop3Port = int.Parse(child.Element("Pop3Port").Value),
                    SmtpServer = child.Element("SmtpServer").Value,
                    SmtpPort = int.Parse(child.Element("SmtpPort").Value),
                    Suffix = child.Element("Suffix").Value,
                    SSL = bool.Parse(child.Element("SSL").Value)
                };
                EmailConfigs.Add(config);
            }
        }
    }
}
