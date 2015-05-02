using EmailCore.POP3;
using EmailCore.SMTP;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            string message = "";
            using (var pop3 = new Receiver("pop3.sina.com", 110))
            {
                pop3.Connect("yant520", "yant923923");
                var msg = pop3.FetchMessage(1);
                message = Receiver.GetMessage(msg);
            }

            //处理回发信息
            var noHtml = RemoveHtmlTags(message);
            var randomStr = GenerateRandomString(message);

            using (var smtp = new Sender("smtp.sina.com"))
            {
                smtp.Connect("yant520", "yant923923");
                var from = new MailAddress("yant520@sina.com", "yant520");
                var tos = new MailAddress[] { 
                    new MailAddress("yant255@163.com","yant255")
                };
                var content = new EmailContent("测试", message);
                smtp.Prepare(from, tos, new MailAddress[] { }, content);
                smtp.Send();
            }
        }

        private static string RemoveHtmlTags(string mailMessage)
        {
            //使用正则表达式来去除html标签
            var partten = @"<\?.+?>";
            return Regex.Replace(mailMessage, partten, "");
        }

        private static string GenerateRandomString(string noHtmlMessage)
        {
            //要获取的随机文本的数量，20个字
            int sendNumber = 20;
            var chars = new char[sendNumber];
            var r = new Random();
            for (int i = 0; i < sendNumber; i++)
            {
                var index = r.Next(noHtmlMessage.Length);
                chars[i] = noHtmlMessage[index];
            }
            return new string(chars);
        }
    }
}
