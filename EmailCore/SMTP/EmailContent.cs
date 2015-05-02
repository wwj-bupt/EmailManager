using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace EmailCore.SMTP
{
    /// <summary>
    /// 邮件内容信息
    /// </summary>
    public class EmailContent
    {
        /// <summary>
        /// 邮件的标题
        /// </summary>
        public string Title { get; private set; }
        /// <summary>
        /// 邮件内容，一般用html格式处理
        /// </summary>
        public string Body { get; private set; }
        /// <summary>
        /// 邮件附件集合
        /// </summary>
        public Attachment[] AttachFile { get; private set; }
        public EmailContent(string title, string body, List<string> attachFileName = null)
        {
            Title = title;
            Body = body;
            AttachFile = GetAttach(attachFileName);
        }

        private Attachment[] GetAttach(List<string> fileNames)
        {
            if (fileNames == null) return null;
            Attachment[] attachs = new Attachment[fileNames.Count];
            for (int i = 0; i < fileNames.Count; i++)
            {
                var attach = new Attachment(fileNames[i]);
                attachs[i] = attach;
            }
            return attachs;
        }
    }
}
