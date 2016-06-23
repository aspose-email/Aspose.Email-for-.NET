using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class GettingMailboxInfo
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "1234.eml";

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

            // Get the size of the mailbox
            long nSize = client.GetMailboxSize();

            Console.WriteLine("Mailbox size is " + nSize + " bytes.");
            // Get mailbox info

            Aspose.Email.Pop3.Pop3MailboxInfo info = client.GetMailboxInfo();

            // Get the number of messages in the mailbox
            int nMessageCount = info.MessageCount;

            Console.WriteLine("Number of messages in mailbox are " + nMessageCount);

            // Get occupied size
            long nOccupiedSize = info.OccupiedSize;

            Console.WriteLine("Occupied size is " + nOccupiedSize);

            Console.WriteLine(Environment.NewLine + "Getting the mailbox information of POP3 server.");
        }
    }
}
