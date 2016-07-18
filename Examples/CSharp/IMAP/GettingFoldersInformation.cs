using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class GettingFoldersInformation
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_IMAP();
            string dstEmail = dataDir + "1234.eml";

            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            // Specify host, username and password for your client
            client.Host = "imap.gmail.com";

            // Set username
            client.Username = "your.username@gmail.com";

            // Set password
            client.Password = "your.password";

            // Set the port to 993. This is the SSL port of IMAP server
            client.Port = 993;

            // Enable SSL
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Get all folders in the currently subscribed folder
                Aspose.Email.Imap.ImapFolderInfoCollection folderInfoColl = client.ListFolders();

                // Iterate through the collection to get folder info one by one
                foreach (Aspose.Email.Imap.ImapFolderInfo folderInfo in folderInfoColl)
                {
                    // Folder name
                    Console.WriteLine("Folder name is " + folderInfo.Name);
                    ImapFolderInfo folderExtInfo = client.GetFolderInfo(folderInfo.Name);
                    // New messages in the folder
                    Console.WriteLine("New message count: " + folderExtInfo.NewMessageCount);
                    // Check whether its readonly
                    Console.WriteLine("Is it readonly? " + folderExtInfo.ReadOnly);
                    // Total number of messages
                    Console.WriteLine("Total number of messages " + folderExtInfo.TotalMessageCount);
                }

                // Disconnect to the remote IMAP server
                client.Dispose();

            }
            catch (Exception ex)
            {
                System.Console.Write(Environment.NewLine + ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Getting folders information from IMAP server.");
        }
    }
}
