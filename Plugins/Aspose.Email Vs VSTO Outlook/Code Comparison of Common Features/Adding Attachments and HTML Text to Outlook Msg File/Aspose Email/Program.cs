using Aspose.Email.Mail;

namespace Aspose_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create an instance of type MailMessage
            MailMessage msg = new MailMessage();

            // Set properties of message like subject, to and HTML body
            // Set subject
            msg.Subject = "This MSG file is created using Aspose.Email for .NET";
            // Set from (sender) address
            msg.Sender = new MailAddress("from@domain.com", "From Name");
            // Set to (recipient) address and name
            msg.To.Add(new MailAddress("to@domain.com", "To Name"));
            // Set HTML body of the email message
            msg.HtmlBody = @"<html><p>This MSG file is created using C# code.</p>" +
                "<p>Microsoft Outlook does not need to be installed on the machine running this code.</p>" +
                "<p>This method is suitable for creating MSG files on the server side.</html>";

            // Add attachments to the message file
            msg.Attachments.Add(new Attachment("image.bmp"));
            msg.Attachments.Add(new Attachment("pic.jpeg"));

            // Save as Outlook MSG file
            string strSaveFile = "TestAspose.msg";
            msg.Save(strSaveFile, MessageFormat.Msg);
        }
    }
}
