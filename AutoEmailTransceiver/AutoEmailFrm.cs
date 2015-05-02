using EmailCore.POP3;
using EmailCore.SMTP;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OpenPop.Mime;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace AutoEmailTransceiver
{
    public partial class AutoEmailFrm : Form
    {
        //定时获取邮件并发送
        private Timer timer;
        //邮件配置信息
        private EmailConfig config;


        //每分钟的秒数
        private const int SecondNumberPerMinute = 60;
        private const int MillisecondNumber = 1000;

        public AutoEmailFrm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer = new Timer();
            timer.Interval = EmailConfigHelper.ServerInterval * SecondNumberPerMinute * MillisecondNumber;
            //初始化可用的邮件服务器列表
            ServerCbx.Items.AddRange(EmailConfigHelper.EmailNames.ToArray());
            ServerCbx.Text = EmailConfigHelper.EmailNames.FirstOrDefault();
        }

        /// <summary>
        /// 判断发送邮件服务是否开启
        /// </summary>
        private bool isStarted = false;

        private void StartBtn_Click(object sender, EventArgs e)
        {
            config = EmailConfigHelper.GetEmailConfig(ServerCbx.Text);
            timer.Tick += (o, arg) =>
            {
                StartTransmit();
            };
            if (isStarted == false)
            {
                StartBtn.Text = "暂停";
                isStarted = true;
                timer.Start();
            }
            else
            {
                StartBtn.Text = "启动";
                isStarted = false;
                timer.Stop();
            }
        }

        //获取未读邮件信息
        private List<OpenPop.Mime.Message> GetUnseenMessages()
        {
            using (var pop3Server = new Receiver(config.Pop3Server, config.Pop3Port, config.SSL))
            {
                pop3Server.Connect(UsernameTxt.Text, PasswordTxt.Text);
                return pop3Server.FetchUnseenMessages();
            }
        }

        /// <summary>
        /// 回发邮件信息
        /// </summary>
        /// <param name="from">发送人用户信息</param>
        /// <param name="to">接收人用户信息</param>
        /// <param name="cc">抄送人用户信息</param>
        /// <param name="content"></param>
        private void SendMessages(MailAddress from, MailAddress[] to, MailAddress[] cc, EmailContent content)
        {
            using (var smtpServer = new Sender(config.SmtpServer, config.SmtpPort))
            {
                smtpServer.Connect(UsernameTxt.Text, PasswordTxt.Text);
                smtpServer.Prepare(from, to, cc, content);
                smtpServer.Send();
            }
        }

        private void StartTransmit()
        {
            var isAutoSend = false;
            string sendMessage = string.Empty;
            if (AutoChk.Checked)
            {
                //随机获取邮件内容，然后回复
                isAutoSend = true;
            }

            if (!isAutoSend && string.IsNullOrEmpty(SendText.Text))
            {
                MessageBox.Show("要发送的邮件文本不能为空");
                return;
            }
            try
            {
                //读取未读邮件集合
                var msgs = GetUnseenMessages();
                //发送人信息
                var from = new MailAddress(UsernameTxt.Text + config.Suffix);
                //遍历集合，并发送邮件
                foreach (var msg in msgs)
                {
                    //指定的要发送的内容
                    sendMessage = isAutoSend ? GetRandomText(Receiver.GetMessage(msg)) : SendText.Text;

                    //接收人信息
                    var to = msg.Headers.From.MailAddress;
                    //发送标题，接收的标题加上系统默认的回复
                    var subject = msg.Headers.Subject + EmailConfigHelper.Subject;
                    var content = new EmailContent(subject, sendMessage);
                    SendMessages(from, new MailAddress[] { to }, new MailAddress[] { }, content);
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }

        private string GetRandomText(string mailMessage)
        {
            //因为收到的邮件很多部分都是html格式的文本，所以获取随机回发的文本需要剔除html标记
            var noHtmlTagMessage = RemoveHtmlTags(mailMessage);
            return GenerateRandomString(noHtmlTagMessage);
        }

        private string RemoveHtmlTags(string mailMessage)
        {
            //使用正则表达式来去除html标签
            var partten = @"<.+?>";
            return Regex.Replace(mailMessage, partten, "");
        }

        private string GenerateRandomString(string noHtmlMessage)
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
