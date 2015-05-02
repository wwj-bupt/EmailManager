using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenPop.Pop3;
using OpenPop.Mime;
using EmailCore.XmlHelper;

namespace EmailCore.POP3
{
    public class Receiver : IDisposable
    {
        private string hostname;
        private int port;
        private bool useSsl;
        private Pop3Client client;
        private SeenIdHelper fileHelper;
        //已读取列表
        private List<string> seenUids;
        /// <summary>
        /// 收件服务器是否连接
        /// </summary>
        public bool IsConnected { get; private set; }

        /// <summary>
        /// 初始化邮件服务器相关信息
        /// </summary>
        /// <param name="hostname">邮件服务器名称，例如：pop3.163.com</param>
        /// <param name="port">服务器端口号，例如：110</param>
        /// <param name="useSsl">是否使用SSL加密连接服务器</param>
        public Receiver(string hostname, int port, bool useSsl = false)
        {
            this.hostname = hostname;
            this.port = port;
            this.useSsl = useSsl;
            IsConnected = false;
            client = new Pop3Client();
            fileHelper = new SeenIdHelper();
            seenUids = fileHelper.Load();
        }

        /// <summary>
        /// 连接邮件服务器
        /// </summary>
        /// <param name="username">用户名，用户名是不带@xxx.com的部分</param>
        /// <param name="password">密码</param>
        public void Connect(string username, string password)
        {
            try
            {
                //连接服务器
                client.Connect(hostname, port, useSsl);
                IsConnected = true;
            }
            catch
            {
                throw new Exception("连接服务器异常");
            }
            try
            {
                //验证用户信息
                client.Authenticate(username, password);
            }
            catch
            {
                throw new Exception("用户名或密码错误");
            }
        }

        /// <summary>
        /// 获取所有邮件列表，包括已经读取的和没有读取的，该方法不会向已读取列表中添加记录
        /// </summary>
        /// <returns></returns>
        public List<Message> FetchAllMessages()
        {
            // 获取邮件数量
            int messageCount = client.GetMessageCount();
            List<Message> allMessages = new List<Message>(messageCount);
            // 遍历所有邮件
            for (int i = messageCount; i > 0; i--)
            {
                allMessages.Add(client.GetMessage(i));
            }
            return allMessages;
        }

        /// <summary>
        /// 获取未读取的邮件列表，获取列表后会向已读取列表中添加记录
        /// </summary>
        /// <param name="hostname">邮件服务器名称，例如：pop3.163.com</param>
        /// <param name="port">服务器端口号，例如：110</param>
        /// <param name="useSsl">是否使用SSL加密连接服务器</param>
        /// <param name="username">服务器登陆用户名</param>
        /// <param name="password">服务器密码</param>
        /// <param name="seenUids">已读邮件列表需要软件自己维护，如设置一个已读列表</param>
        /// <returns></returns>
        public List<Message> FetchUnseenMessages()
        {
            // 获取邮件数量
            int messageCount = client.GetMessageCount();

            List<Message> newMessages = new List<Message>();
            // 遍历所有邮件列表
            for (int i = 1; i <= messageCount; i++)
            {
                // 如果邮件不存在于已读列表，则获取该邮件，并更新已读邮件列表
                Message unseenMessage = client.GetMessage(i);
                if (!seenUids.Contains(unseenMessage.Headers.MessageId))
                {
                    newMessages.Add(unseenMessage);
                    seenUids.Add(unseenMessage.Headers.MessageId);
                }
            }
            return newMessages;
        }

        /// <summary>
        /// 根据邮件编号获取邮件信息，读取邮件后会向已读取列表中添加记录
        /// </summary>
        /// <param name="messageNumber">邮件的编号，从1开始、到邮件列表的数量</param>
        /// <returns></returns>
        public Message FetchMessage(int messageNumber)
        {
            var msg = client.GetMessage(messageNumber);
            if (!seenUids.Contains(msg.Headers.MessageId))
            {
                seenUids.Add(msg.Headers.MessageId);
            }
            return msg;
        }

        /// <summary>
        /// 获取邮箱信息
        /// </summary>
        /// <param name="msg">要读取的邮件信息</param>
        /// <returns></returns>
        public static string GetMessage(Message msg)
        {
            var message = GetMessageHtml(msg);
            if (message != null) return message;
            message = GetMessageText(msg);
            if (message != null) return message;
            return null;
        }

        /// <summary>
        /// 将邮件信息的内容读取为文本
        /// </summary>
        /// <param name="msg">要读取的邮件信息</param>
        /// <returns></returns>
        public static string GetMessageText(Message msg)
        {
            var messagePart = msg.FindFirstPlainTextVersion();
            if (messagePart == null) return null;
            return messagePart.GetBodyAsText();
        }

        /// <summary>
        /// 将邮件信息的内容读取为html文本
        /// </summary>
        /// <param name="msg">要读取的邮件信息</param>
        /// <returns></returns>
        public static string GetMessageHtml(Message msg)
        {
            var messagePart = msg.FindFirstHtmlVersion();
            if (messagePart == null) return null;
            return messagePart.GetBodyAsText();
        }

        /// <summary>
        /// 通过邮件的MessageId删除邮件
        /// </summary>
        /// <param name="client">已连接并验证的POP3对象</param>
        /// <param name="messageId">MessageID。 指向<see cref="MessageHeader.MessageId"/></param>
        /// <returns><see langword="true"/> 邮件被删除 <see langword="false"/> 其他</returns>
        public bool DeleteMessageByMessageId(string messageId)
        {
            // 获取邮件数量
            int messageCount = client.GetMessageCount();
            for (int messageItem = messageCount; messageItem > 0; messageItem--)
            {
                // 如果当前的邮件MessageId和参数提供的相同，则删除
                if (client.GetMessageHeaders(messageItem).MessageId == messageId)
                {
                    client.DeleteMessage(messageItem);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 释放对象
        /// </summary>
        public void Dispose()
        {
            fileHelper.Save(seenUids);
            client.Dispose();
        }
    }
}
