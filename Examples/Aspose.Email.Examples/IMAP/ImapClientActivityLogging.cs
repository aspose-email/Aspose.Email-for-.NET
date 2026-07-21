using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.IMAP
{
    class ImapClientActivityLogging
    {
        public static void Run()
        {
            ImapClient client = new ImapClient("imap.gmail.com", 993, "user@gmail.com", "password");

            // Set security mode
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Get the message info collection
                ImapMessageInfoCollection list = client.ListMessages();

                // Download each message
                for (int i = 0; i < list.Count; i++)
                {
                    // Save the EML file locally
                    client.SaveMessage(list[i].UniqueId, Data.Out + list[i].UniqueId + ".eml");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }
    }
}
