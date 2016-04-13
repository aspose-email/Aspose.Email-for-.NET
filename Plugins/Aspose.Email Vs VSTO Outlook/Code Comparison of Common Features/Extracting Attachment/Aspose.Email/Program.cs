using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Email
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"E:\Aspose\Aspose VS VSTO\Sample Files\ExtractAttachment.msg";
            string AttachmentFilePath = @"E:\Aspose\Aspose VS VSTO\Sample Files\";

            //Create an instance of MailMessage and load an email file
            MailMessage mailMsg = MailMessage.Load(FilePath, MailMessageLoadOptions.DefaultEml);

            foreach (Attachment attachment in mailMsg.Attachments)
            {
                //To display the the attachment file name
                attachment.Save(AttachmentFilePath+attachment.Name);
                Console.WriteLine(attachment.Name);
            }

            mailMsg.From = "sender@sender.com";
            mailMsg.To.Add("receiver@receiver.com");

            SmtpClient client = new SmtpClient("smtp.server.com");
            client.Port = 25;
            client.Username = "UserName";
            client.Password = "password";
            {
                client.Send(mailMsg);
            }
        }
    }
}
