using Aspose.Email.Mail;
namespace Aspose_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create an instance of the Aspose.Email.MailMessage class
            MailMessage msg = new MailMessage();

            // Set recipients information
            msg.To = "to@domain.com";
            msg.CC = "cc@domain.com";

            // Set the subject
            msg.Subject = "Subject";

            // Set HTML body
            msg.HtmlBody = "<h3>HTML Heading 3</h3> <u>This is underlined text</u>";

            // Add an attachment
            msg.Attachments.Add(new Attachment("image.bmp"));

            // Save it in local disk
            string strMsg = "TestAspose.msg";
            msg.Save(strMsg, MessageFormat.Msg);
        }
    }
}
