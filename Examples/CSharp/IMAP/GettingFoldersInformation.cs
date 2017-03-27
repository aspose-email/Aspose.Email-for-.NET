using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class GettingFoldersInformation
    {
        public static void Run()
        {            
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
                // ExStart:GettingFoldersInformation
                // Get all folders in the currently subscribed folder
                ImapFolderInfoCollection folderInfoColl = client.ListFolders();

                // Iterate through the collection to get folder info one by one
                foreach (ImapFolderInfo folderInfo in folderInfoColl)
                {
                    // Folder name and get New messages in the folder
                    Console.WriteLine("Folder name is " + folderInfo.Name);
                    ImapFolderInfo folderExtInfo = client.GetFolderInfo(folderInfo.Name);
                    Console.WriteLine("New message count: " + folderExtInfo.NewMessageCount);
                    Console.WriteLine("Is it readonly? " + folderExtInfo.ReadOnly);
                    Console.WriteLine("Total number of messages " + folderExtInfo.TotalMessageCount);
                }
                // ExEnd:GettingFoldersInformation

                // Disconnect to the remote IMAP server
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            Console.WriteLine(Environment.NewLine + "Getting folders information from IMAP server.");
        }
    }
}