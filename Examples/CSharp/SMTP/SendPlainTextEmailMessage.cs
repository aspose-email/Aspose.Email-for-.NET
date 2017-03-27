using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class SendPlainTextEmailMessage
    {
        public static void Run()
        {
            // ExStart:SendPlainTextEmailMessage
            //Create an instance of the MailMessage class
            MailMessage message = new MailMessage();

            // Set From field, To field and Plain text body
            message.From = "sender@sender.com";
            message.To.Add("receiver@receiver.com");
            message.Body = "This is Plain Text Body";

            // Create an instance of the SmtpClient class
            SmtpClient client = new SmtpClient();

            // And Specify your mailing host server, Username, Password and Port
            client.Host = "smtp.server.com";
            client.Username = "Username";
            client.Password = "Password";
            client.Port = 25;

            try
            {
                //Client.Send will send this message
                client.Send(message);
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }
            // ExEnd:SendPlainTextEmailMessage
        }
    }
}