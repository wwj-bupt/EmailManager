using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoEmailTransceiver
{
    /// <summary>
    /// 用于获取邮件配置的模型对象
    /// </summary>
    public class EmailConfig
    {
        public string Name { get; set; }
        /// <summary>
        /// Pop3服务器地址
        /// </summary>
        public string Pop3Server { get; set; }
        /// <summary>
        /// Pop3服务器端口，默认为110
        /// </summary>
        public int Pop3Port { get; set; }
        /// <summary>
        /// Stmp服务器地址
        /// </summary>
        public string SmtpServer { get; set; }
        /// <summary>
        /// Smtp服务器断开，默认为25
        /// </summary>
        public int SmtpPort { get; set; }
        public string Suffix { get; set; }
        public bool SSL { get; set; }

        public EmailConfig()
        {
            Pop3Port = 110;
            SmtpPort = 25;
        }
    }
}
