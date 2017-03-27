using Aspose.Email.Clients.Smtp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class SendEmailMessage
    {
        public static void Run()
        {
            try
            {
                string dataDir = RunExamples.GetDataDir_KnowledgeBase();
                SmtpClient client = new SmtpClient("smtp.gmail.com");

                // Set username, Password, Port No, and SecurityOptions
                client.Username = "user@gmail.com";
                client.Password = "your.password";
                client.Port = 587;

                // ExStart:SendEmailMessage
                MailMessage msg = new MailMessage();

                // Set From, To whom
                msg.From = new MailAddress("user@gmail.com");
                msg.To.Add(new MailAddress("user@gmail.com"));

                // Set the subject, priority, Set the message bodies (plain text and HTML), and Add the attachment
                msg.Subject = "subject";
                msg.Priority = MailPriority.High;
                msg.Body = "Please open with html browser";
                msg.HtmlBody = "<bold>this is the mail body</bold>";
                Attachment attachment = new Attachment(dataDir + "File.txt");
                msg.Attachments.Add(attachment);
                client.Send(msg);
                // ExEnd:SendEmailMessage
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
