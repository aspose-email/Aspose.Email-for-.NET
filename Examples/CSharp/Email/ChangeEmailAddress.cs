using Aspose.Email.Mail;
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

namespace Aspose.Email.Examples.CSharp.Email
{
    class ChangeEmailAddress
    {
        public static void Run()
        {
            // ExStart:ChangeEmailAddress
            MailMessage message = new MailMessage("Sender Name <sender@from.com>", "Kyle Huang <kyle@to.com>");

            // A To address with a friendly name can also be specified like this
            message.To.Add(new MailAddress("kyle@to.com", "Kyle Huang"));

            // Specify Cc and Bcc email address along with a friendly name
            message.CC.Add(new MailAddress("guangzhou@cc.com", "Guangzhou Team"));
            message.Bcc.Add(new MailAddress("ahaq@bcc.com", "Ammad ulHaq "));

            // Create an instance of SmtpClient Class and Specify your mailing host server, Username, Password, Port
            SmtpClient client = new SmtpClient();
            client.Host = "smtp.server.com";
            client.Username = "Username";
            client.Password = "Password";
            client.Port = 25;

            try
            {
                // Client.Send will send this message
                client.Send(message);
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }
            // ExEnd:ChangeEmailAddress
        }
    }
}
