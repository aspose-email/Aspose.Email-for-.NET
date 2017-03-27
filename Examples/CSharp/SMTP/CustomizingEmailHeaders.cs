using System;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class CustomizingEmailHeaders
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "MsgHeaders.msg";

            // Create an instance MailMessage class
            MailMessage msg = new MailMessage();

            // Specify ReplyTo
            msg.ReplyToList.Add("reply@reply.com");

            // From field
            msg.From = "sender@sender.com";

            // To field
            msg.To.Add("receiver1@receiver.com");

            // Adding Cc and Bcc Addresses
            msg.CC.Add("receiver2@receiver.com");
            msg.Bcc.Add("receiver3@receiver.com");

            // Message subject
            msg.Subject = "test mail";

            // Specify Date
            msg.Date = new DateTime(2006, 3, 6);

            // Specify XMailer
            msg.XMailer = "Aspose.Email";

            // Specify Secret Header
            msg.Headers.Add("secret-header", "mystery");

            // Save message to disc
            msg.Save(dstEmail, SaveOptions.DefaultMsgUnicode);

            Console.WriteLine(Environment.NewLine + "Message saved with customized headers successfully at " + dstEmail);
        }
    }
}
