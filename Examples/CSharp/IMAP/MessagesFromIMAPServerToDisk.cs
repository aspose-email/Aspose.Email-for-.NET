using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class MessagesFromIMAPServerToDisk
    {
        public static void Run()
        {
            // ExStart:MessagesFromIMAPServerToDisk
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
                // Select the inbox folder and Get the message info collection
                client.SelectFolder(ImapFolderInfo.InBox);
                ImapMessageInfoCollection list = client.ListMessages();

                // Download each message
                for (int i = 0; i < list.Count; i++)
                {
                    // Save the EML file locally
                    client.SaveMessage(list[i].UniqueId, dataDir + list[i].UniqueId + ".eml");
                }
                // Disconnect to the remote IMAP server
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            // ExStart:MessagesFromIMAPServerToDisk
            Console.WriteLine(Environment.NewLine + "Downloaded messages from IMAP server.");
        }
    }
}
