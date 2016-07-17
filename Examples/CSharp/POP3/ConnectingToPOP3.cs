using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class ConnectingToPOP3
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "1234.eml";

            // Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();

            // Specify host, username and password for your client
            client.Host = "pop.gmail.com";

            // Set username
            client.Username = "your.username@gmail.com";

            // Set password
            client.Password = "your.password";

            // Set the port to 995. This is the SSL port of POP3 server
            client.Port = 995;

            // Enable SSL
            client.SecurityOptions = SecurityOptions.Auto;

            Console.WriteLine(Environment.NewLine + "Connected to POP3 server.");
        }
    }
}
