using System;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.SMTP
{
    class ExportAsEML
    {
        public static void Run()
        {
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
            message.Save(Data.Out + "Message.eml", SaveOptions.DefaultEml);  

            Console.WriteLine(Environment.NewLine + $"Email saved at {Data.Out}/Message.eml");
        }
    }
}
