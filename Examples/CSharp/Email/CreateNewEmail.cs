 

using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;
using System.Configuration;
using System.Data;
using Aspose.Email.Mail.Bounce;

namespace Aspose.Email.Examples.CSharp.Email
{
    class CreateNewEmail
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            // Create a new instance of MailMessage class
            MailMessage message = new MailMessage();

            // Set subject of the message
            message.Subject = "New message created by Aspose.Email for .NET";

            // Set Html body
            message.IsBodyHtml = true;
            message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/><font color=blue>This line is in blue color</font>";

            // Set sender information
            message.From = "from@domain.com";

            // Add TO recipients
            message.To.Add("to1@domain.com");
            message.To.Add("to2@domain.com");

            // Add CC recipients
            message.CC.Add("cc1@domain.com");
            message.CC.Add("cc2@domain.com");

            // Save message in EML, MSG and MHTML formats
            message.Save(dataDir + "Message.eml", Aspose.Email.Mail.SaveOptions.DefaultEml);
            message.Save(dataDir + "Message.msg", Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode);
            message.Save(dataDir + "Message.mhtml", Aspose.Email.Mail.SaveOptions.DefaultMhtml);

            Console.WriteLine(Environment.NewLine + "Created new email in EML, MSG and MHTML formats successfully.");
        }
    }
}
