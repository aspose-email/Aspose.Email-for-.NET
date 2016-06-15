using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.POP3
{
    class RetrievingEmailMessages
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "message.msg";

            //Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();

            //Specify host, username and password for your client
            client.Host = "pop.gmail.com";

            // Set username
            client.Username = "your.username@gmail.com";

            // Set password
            client.Password = "your.password";

            // Set the port to 995. This is the SSL port of POP3 server
            client.Port = 995;

            // Enable SSL
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                int messageCount = client.GetMessageCount();
                // Create an instance of the MailMessage class
                MailMessage msg;

                for (int i = 1; i <= messageCount; i++)
                {
                    //Retrieve the message in a MailMessage class
                    msg = client.FetchMessage(i);

                    Console.WriteLine("From:" + msg.From.ToString());
                    Console.WriteLine("Subject:" + msg.Subject);
                    Console.WriteLine(msg.HtmlBody);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                client.Disconnect();
            } 

            Console.WriteLine(Environment.NewLine + "Retrieved email messages using POP3 ");
        }
    }
}
