using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.POP3
{
    class Pop3ClientActivityLogging
    {
        public static void Run()
        {
            Pop3Client client = new Pop3Client("pop.gmail.com", 995, "user@gmail.com", "password");

            // Set security mode
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Get the message info collection
                Pop3MessageInfoCollection list = client.ListMessages();

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
