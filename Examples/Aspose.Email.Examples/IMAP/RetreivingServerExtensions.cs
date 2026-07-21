using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    class RetreivingServerExtensions
    {
        public static void Run()
        {
            // Connect and log in to IMAP
            ImapClient client = new ImapClient("imap.gmail.com", "username", "password");
            string[] getCapabilities = client.GetCapabilities();
            foreach (string getCap in getCapabilities)
            {
                Console.WriteLine(getCap);
            }
        }
    }
}