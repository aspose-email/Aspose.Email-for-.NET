using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class SendMessageAsTNEF
    {
        public static void Run()
        {
            try
            {
                // ExStart:SendMessageAsTNEF
                var emlFileName = RunExamples.GetDataDir_Email() + "Message.eml";     // A TNEF Email

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
                // ExEnd:SendMessageAsTNEF
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}
