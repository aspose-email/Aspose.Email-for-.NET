using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class SendingBulkEmails
    {
       
        public static void Run()
        {
            // ExStart:SendingBulkEmails
            // Create SmtpClient as client and specify server, port, user name and password
            SmtpClient client = new SmtpClient("mail.server.com", 25, "Username", "Password");

            // Create instances of MailMessage class and Specify To, From, Subject and Message
            MailMessage message1 = new MailMessage("msg1@from.com", "msg1@to.com", "Subject1", "message1, how are you?");
            MailMessage message2 = new MailMessage("msg1@from.com", "msg2@to.com", "Subject2", "message2, how are you?");
            MailMessage message3 = new MailMessage("msg1@from.com", "msg3@to.com", "Subject3", "message3, how are you?");

            // Create an instance of MailMessageCollection class
            MailMessageCollection manyMsg = new MailMessageCollection();
            manyMsg.Add(message1);
            manyMsg.Add(message2);
            manyMsg.Add(message3);

            // Use client.BulkSend function to complete the bulk send task
            try
            {
                // Send Message using BulkSend method
                client.Send(manyMsg);                
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
            // ExEnd:SendingBulkEmails
        }
       
    }
}
