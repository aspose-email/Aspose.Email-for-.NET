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
    class ConnectingGmailSMTPServer
    {
        public static void Run()
        {
            try
            {
                // ExStart:ConnectingGmailSMTPServer
                SmtpClient client = new SmtpClient("smtp.gmail.com");

                // Set username, Password, Port No, and SecurityOptions
                client.Username = "your.email@gmail.com";
                client.Password = "your.password";
                client.Port = 587;
                // ExEnd:ConnectingGmailSMTPServer

                // Declare message as MailMessage instance
                MailMessage message = new MailMessage();

                // Use MailMessage properties like specify sender, recipient and message
                message.To = "newcustomeronnet@gmail.com";
                message.From = "newcustomeronnet@gmail.com";
                message.Subject = "Test Email";
                message.Body = "Hello World!";

                // Client will send this message
                client.Send(message);
                Console.WriteLine("Message sent");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
