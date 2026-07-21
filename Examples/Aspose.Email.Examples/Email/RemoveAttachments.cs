// Demonstrates how to add attachments to a MailMessage, remove one of them,
// save the result, and list the remaining attachments.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class RemoveAttachments
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                From = "sender@sender.com",
                To = "receiver@gmail.com"
            };

            var attachment = new Attachment(Data.Email/"1.txt");
            eml.Attachments.Add(attachment);
            eml.Attachments.Add(new Attachment(Data.Email/"1.jpg"));
            eml.Attachments.Add(new Attachment(Data.Email/"1.doc"));
            eml.Attachments.Add(new Attachment(Data.Email/"1.rar"));
            eml.Attachments.Add(new Attachment(Data.Email/"1.pdf"));

            eml.Attachments.Remove(attachment);
            eml.Save(Data.Out/"RemoveAttachments.msg", SaveOptions.DefaultMsgUnicode);

            foreach (var remaining in eml.Attachments)
                Console.WriteLine(remaining.Name);
        }
    }
}
