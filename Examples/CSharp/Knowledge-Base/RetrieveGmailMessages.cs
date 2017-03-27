using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;
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
    class RetrieveGmailMessages
    {
        public static void Run()
        {
            try
            {
                // Create a POP3 client
                Pop3Client client;
                client = new Pop3Client();

                // Basic settings (required)
                client.Host = "pop.gmail.com";
                client.Username = "user@gmail.com";
                client.Password = "pwd";
                client.Port = 995;
                client.SecurityOptions = SecurityOptions.SSLImplicit;

                // ExStart:RetrieveGmailMessages
                // Get the number of messages in the mailbox
                int messageCount = client.GetMessageCount();

                // Iterate through the messages and retrieve one by one
                for (int i = 1; i <= messageCount; i++)
                {
                    // Create an object of type MailMessage and Retrieve the message in MailMessage format directly
                    Aspose.Email.MailMessage msg;
                    msg = client.FetchMessage(i);
                    Console.WriteLine("From:" + msg.From.ToString());
                    Console.WriteLine("Subject:" + msg.Subject);
                    Console.WriteLine(msg.HtmlBody);
                    msg.Save(i + "Getmessage_out.eml");
                }
                // ExEnd:RetrieveGmailMessages
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
