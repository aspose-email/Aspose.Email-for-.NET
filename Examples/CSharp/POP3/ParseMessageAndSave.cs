using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;

namespace Aspose.Email.Examples.CSharp.POP3
{
    class ParseMessageAndSave
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "first-message.eml";

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
                //Fetch the message by its sequence number
                MailMessage msg = client.FetchMessage(1);

                //Save the message using its subject as the file name
                msg.Save(dstEmail, Aspose.Email.Mail.SaveOptions.DefaultEml);
                client.Disconnect();

            }
            catch (Exception ex)
            {
                Console.WriteLine(Environment.NewLine + ex.Message);
            }
            finally
            {
                client.Disconnect();
            } 

            Console.WriteLine(Environment.NewLine + "Downloaded email using POP3. Message saved at " + dstEmail);
        }
    }
}
