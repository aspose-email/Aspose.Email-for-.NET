using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class RetrievingEmailMessages
    {
        public static void Run()
        {
            //ExStart:RetrievingEmailMessages
            // Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();

            // Specify host, username, password, Port and SecurityOptions for your client
            client.Host = "pop.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 995;
            client.SecurityOptions = SecurityOptions.Auto;
            try
            {
                int messageCount = client.GetMessageCount();
                // Create an instance of the MailMessage class and Retrieve message
                MailMessage message;
                for (int i = 1; i <= messageCount; i++)
                {
                    message = client.FetchMessage(i);
                    Console.WriteLine("From:" + message.From);
                    Console.WriteLine("Subject:" + message.Subject);
                    Console.WriteLine(message.HtmlBody);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                client.Dispose();
            }
            //ExEnd:RetrievingEmailMessages
            Console.WriteLine(Environment.NewLine + "Retrieved email messages using POP3 ");
        }
    }
}
