using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    class SendMessageAsTNEF
    {
        public static void Run()
        {
            try
            {
                var emlFileName = Data.Email + "Message.eml";     // A TNEF Email

                // Load from eml
                MailMessage eml1 = MailMessage.Load(emlFileName, new EmlLoadOptions());
                eml1.From = "somename@gmail.com";
                eml1.To.Clear();
                eml1.To.Add(new MailAddress("first.last@test.com"));
                eml1.Subject = "With PreserveTnef flag during loading";
                eml1.Date = DateTime.Now;
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "somename", "password");
                client.SecurityOptions = SecurityOptions.Auto;
                client.UseTnef = true;     // Use this flag to send as TNEF
                client.Send(eml1);
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}
