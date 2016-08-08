using System;
using Aspose.Email.Mail;

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
            // ExStart:SendMessageAsTNEF
            var emlFileName = RunExamples.GetDataDir_Email() +  "Message.eml";     // A TNEF Email
            MailMessageLoadOptions options = new MailMessageLoadOptions();
            options.MessageFormat = MessageFormat.Eml;         
            MailMessage eml = MailMessage.Load(emlFileName, options);
            eml.From = "somename@gmail.com";
            eml.To.Clear();
            eml.To.Add(new MailAddress("first.last@test.com"));
            eml.Subject = "With PreserveTnef flag during loading";
            eml.Date = DateTime.Now;

            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "somename", "password");
            client.SecurityOptions = SecurityOptions.Auto;
            client.UseTnef = true;     // Use this flag to send as TNEF
            client.Send(eml);
            // ExEnd:SendMessageAsTNEF
        }
        
    }
}
