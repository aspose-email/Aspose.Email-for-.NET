using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class GettingMailboxInfo
    {
        public static void Run()
        {
            // Create an instance of the Pop3Client class
            //ExStart:GettingMailboxInfo
            Pop3Client client = new Pop3Client();
            // Specify host, username, password, Port and SecurityOptions for your client
            client.Host = "pop.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 995;
            client.SecurityOptions = SecurityOptions.Auto;

            // Get the size of the mailbox,  Get mailbox info, number of messages in the mailbox
            long nSize = client.GetMailboxSize();
            Console.WriteLine("Mailbox size is " + nSize + " bytes.");
            Pop3MailboxInfo info = client.GetMailboxInfo();
            int nMessageCount = info.MessageCount;
            Console.WriteLine("Number of messages in mailbox are " + nMessageCount);
            long nOccupiedSize = info.OccupiedSize;
            Console.WriteLine("Occupied size is " + nOccupiedSize);
            //ExEnd:GettingMailboxInfo
            Console.WriteLine(Environment.NewLine + "Getting the mailbox information of POP3 server.");
        }
    }
}
