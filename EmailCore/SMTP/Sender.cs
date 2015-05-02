using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace EmailCore.SMTP
{
    public class Sender : IDisposable
    {
        private string host;
        private int port;
        private MailMessage mail;
        private SmtpClient client;
        /// <summary>
        /// 发件服务器是否连接
        /// </summary>
        public bool IsConnected { get; private set; }

        /// <summary>
        /// 初始化邮件发送服务
        /// </summary>
        /// <param name="host">要发送的邮箱主机</param>
        public Sender(string host, int port = 25)
        {
            this.host = host;
            this.port = port;
            client = new SmtpClient();
            IsConnected = true;
        }

        /// <summary>
        /// 连接到邮件服务器
        /// </summary>
        /// <param name="username">邮件的用户名</param>
        /// <param name="password">邮件的密码</param>
        public void Connect(string username, string password)
        {
            client.Host = host;
            client.Port = port;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential(username, password);
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
        }

        /// <summary>
        /// 准备要发送的相关信息
        /// </summary>
        /// <param name="from">发信人信息</param>
        /// <param name="to">收信人信息，可以多个收信人</param>
        /// <param name="cc">抄送人信息，可以抄送多个</param>
        /// <param name="content">邮件内容信息</param>
        public void Prepare(MailAddress from, MailAddress[] to, MailAddress[] cc, EmailContent content)
        {
            mail = new MailMessage();
            mail.From = from;
            foreach (var receiver in to)
            {
                mail.To.Add(receiver);
            }
            foreach (var copy in cc)
            {
                mail.CC.Add(copy);
            }

            mail.Subject = content.Title;
            mail.IsBodyHtml = true;
            mail.Body = content.Body;
            mail.Priority = MailPriority.Normal;
            mail.BodyEncoding = Encoding.UTF8;
            mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
            if (content.AttachFile == null) return;
            foreach (var att in content.AttachFile)
            {
                mail.Attachments.Add(att);
            }
        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <exception cref="InvalidOperationException">没有指定服务器、发送人或收件人</exception>
        public void Send()
        {
            try
            {
                client.Send(mail);
            }
            catch (InvalidOperationException ioExp)
            {
                throw ioExp;
            }
        }

        /// <summary>
        /// 释放对象资源
        /// </summary>
        public void Dispose()
        {
            client.Dispose();
        }
    }
}
