using System;
using System.IO;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ReadMessagesRecursively
    {
        // ExStart:ReadMessagesRecursively
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
                // The root folder (which will be created on disk) consists of host and username
                string rootFolder = client.Host + "-" + client.Username;

                // Create the root folder and List all the folders from IMAP server
                Directory.CreateDirectory(rootFolder);
                ImapFolderInfoCollection folderInfoCollection = client.ListFolders();
                foreach (ImapFolderInfo folderInfo in folderInfoCollection)
                {
                    // Call the recursive method to read messages and get sub-folders
                    ListMessagesInFolder(folderInfo, rootFolder, client);
                }
                // Disconnect to the remote IMAP server
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }

            Console.WriteLine(Environment.NewLine + "Downloaded messages recursively from IMAP server.");
        }

        /// Recursive method to get messages from folders and sub-folders
        private static void ListMessagesInFolder(ImapFolderInfo folderInfo, string rootFolder, ImapClient client)
        {
            // Create the folder in disk (same name as on IMAP server)
            string currentFolder = RunExamples.GetDataDir_IMAP();
            Directory.CreateDirectory(currentFolder);

            // Read the messages from the current folder, if it is selectable
            if (folderInfo.Selectable)
            {
                // Send status command to get folder info
                ImapFolderInfo folderInfoStatus = client.GetFolderInfo(folderInfo.Name);
                Console.WriteLine(folderInfoStatus.Name + " folder selected. New messages: " + folderInfoStatus.NewMessageCount + ", Total messages: " + folderInfoStatus.TotalMessageCount);

                // Select the current folder and List messages
                client.SelectFolder(folderInfo.Name);
                ImapMessageInfoCollection msgInfoColl = client.ListMessages();
                Console.WriteLine("Listing messages....");
                foreach (ImapMessageInfo msgInfo in msgInfoColl)
                {
                    // Get subject and other properties of the message
                    Console.WriteLine("Subject: " + msgInfo.Subject);
                    Console.WriteLine("Read: " + msgInfo.IsRead + ", Recent: " + msgInfo.Recent + ", Answered: " + msgInfo.Answered);

                    // Get rid of characters like ? and :, which should not be included in a file name and Save the message in MSG format
                    string fileName = msgInfo.Subject.Replace(":", " ").Replace("?", " ");
                    MailMessage msg = client.FetchMessage(msgInfo.SequenceNumber);
                    msg.Save(currentFolder + "\\" + fileName + "-" + msgInfo.SequenceNumber + ".msg", SaveOptions.DefaultMsgUnicode);
                }
                Console.WriteLine("============================\n");
            }
            else
            {
                Console.WriteLine(folderInfo.Name + " is not selectable.");
            }

            try
            {
                // If this folder has sub-folders, call this method recursively to get messages
                ImapFolderInfoCollection folderInfoCollection = client.ListFolders(folderInfo.Name);
                foreach (ImapFolderInfo subfolderInfo in folderInfoCollection)
                {
                    ListMessagesInFolder(subfolderInfo, rootFolder, client);
                }
            }
            catch (Exception) { }            
        }
        // ExEnd:ReadMessagesRecursively
    }
}
