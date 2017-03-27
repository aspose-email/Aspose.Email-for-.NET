using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class ParseMessageAndSave
    {
        public static void Run()
        {
            // ExStart:ParseMessageAndSave
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_POP3();

            // Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();

            // Specify host, username and password, Port and SecurityOptions for your client
            client.Host = "pop.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 995;
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Fetch the message by its sequence number and Save the message using its subject as the file name
                MailMessage msg = client.FetchMessage(1);
                msg.Save(dataDir + "first-message_out.eml", SaveOptions.DefaultEml);
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(Environment.NewLine + ex.Message);
            }
            finally
            {
                client.Dispose();
            }
            // ExEnd:ParseMessageAndSave
            Console.WriteLine(Environment.NewLine + "Downloaded email using POP3. Message saved at " + dataDir + "first-message_out.eml");
        }
    }
}
