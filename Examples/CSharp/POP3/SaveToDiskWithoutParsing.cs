using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class SaveToDiskWithoutParsing
    {
        public static void Run()
        {
            //ExStart:SaveToDiskWithoutParsing
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "InsertHeaders.eml";

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
                // Save message to disk by message sequence number
                client.SaveMessage(1, dstEmail);
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.WriteLine(Environment.NewLine + "Retrieved email messages using POP3 ");
            //ExEnd:SaveToDiskWithoutParsing
        }
    }
}
