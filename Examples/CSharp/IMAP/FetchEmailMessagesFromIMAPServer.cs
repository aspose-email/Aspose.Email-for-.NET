using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class FetchEmailMessagesFromIMAPServer
    {
        public static void Run()
        {            
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_IMAP();

            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            // Specify host, username, password, Port and SecurityOptions for your client
            client.Host = "imap.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 993;
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // ExStart:FetchEmailMessagesFromIMAPServer
                // Select the inbox folder and Get the message info collection
                client.SelectFolder(ImapFolderInfo.InBox);
                ImapMessageInfoCollection list = client.ListMessages();

                // Download each message
                for (int i = 0; i < list.Count; i++)
                {
                    // Save the EML file locally
                    client.SaveMessage(list[i].UniqueId, dataDir + list[i].UniqueId + ".eml");
                }
                // ExEnd:FetchEmailMessagesFromIMAPServer

                // Disconnect to the remote IMAP server
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }            
        }
    }
}
