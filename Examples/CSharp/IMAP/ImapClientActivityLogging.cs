using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ImapClientActivityLogging
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_IMAP();

            // ExStart:ImapClientActivityLogging
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
                    client.SaveMessage(list[i].UniqueId, dataDir + list[i].UniqueId + ".eml");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
            // ExEnd:ImapClientActivityLogging
        }
    }
}
