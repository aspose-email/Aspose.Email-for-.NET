using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.SMTP
{
    class CustomizingEmailHeaders
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "MsgHeaders.msg";

            //Create an instance MailMessage class
            MailMessage msg = new MailMessage();

            //Specify ReplyTo
            msg.ReplyToList.Add("reply@reply.com");

            //From field
            msg.From = "sender@sender.com";

            //To field
            msg.To.Add("receiver1@receiver.com");

            //Adding Cc and Bcc Addresses
            msg.CC.Add("receiver2@receiver.com");
            msg.Bcc.Add("receiver3@receiver.com");

            //Message subject
            msg.Subject = "test mail";

            //Specify Date
            msg.Date = new System.DateTime(2006, 3, 6);

            //Specify XMailer
            msg.XMailer = "Aspose.Email";

            //Specify Secret Header
            msg.Headers.Add("secret-header", "mystery");

            //Save message to disc
            msg.Save(dstEmail, Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode);

            Console.WriteLine(Environment.NewLine + "Message saved with customized headers successfully at " + dstEmail);
        }
    }
}
