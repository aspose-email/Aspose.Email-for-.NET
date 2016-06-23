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
    class ReadMessagesRecursively
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_IMAP();
            string dstEmail = dataDir + "1234.eml";

            //Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            //Specify host, username and password for your client
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
                // The root folder (which will be created on disk) consists of host and username
                string rootFolder = client.Host + "-" + client.Username;
                // Create the root folder
                Directory.CreateDirectory(rootFolder);

                // List all the folders from IMAP server
                ImapFolderInfoCollection folderInfoCollection = client.ListFolders();
                foreach (ImapFolderInfo folderInfo in folderInfoCollection)
                {
                    // Call the recursive method to read messages and get sub-folders
                    ListMessagesInFolder(folderInfo, rootFolder, client);
                }


                //Disconnect to the remote IMAP server
                client.Dispose();

            }
            catch (Exception ex)
            {
                System.Console.Write(Environment.NewLine + ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Downloaded messages recursively from IMAP server.");
        }

        /// <summary>
        /// Recursive method to get messages from folders and sub-folders
        /// </summary>
        /// <param name="folderInfo"></param>
        /// <param name="rootFolder"></param>
        private static void ListMessagesInFolder(ImapFolderInfo folderInfo, string rootFolder, ImapClient client)
        {
            // Create the folder in disk (same name as on IMAP server)
            string currentFolder = Path.Combine(Path.GetFullPath("../../../Data/"), folderInfo.Name);
            Directory.CreateDirectory(currentFolder);

            // Read the messages from the current folder, if it is selectable
            if (folderInfo.Selectable == true)
            {
                // Send status command to get folder info
                ImapFolderInfo folderInfoStatus = client.GetFolderInfo(folderInfo.Name);
                Console.WriteLine(folderInfoStatus.Name + " folder selected. New messages: " + folderInfoStatus.NewMessageCount +
                            ", Total messages: " + folderInfoStatus.TotalMessageCount);

                // Select the current folder
                client.SelectFolder(folderInfo.Name);
                // List messages
                ImapMessageInfoCollection msgInfoColl = client.ListMessages();
                Console.WriteLine("Listing messages....");
                foreach (ImapMessageInfo msgInfo in msgInfoColl)
                {
                    // Get subject and other properties of the message
                    Console.WriteLine("Subject: " + msgInfo.Subject);
                    Console.WriteLine("Read: " + msgInfo.IsRead + ", Recent: " + msgInfo.Recent + ", Answered: " + msgInfo.Answered);

                    // Get rid of characters like ? and :, which should not be included in a file name
                    string fileName = msgInfo.Subject.Replace(":", " ").Replace("?", " ");

                    // Save the message in MSG format
                    MailMessage msg = client.FetchMessage(msgInfo.SequenceNumber);
                    msg.Save(currentFolder + "\\" + fileName + "-" + msgInfo.SequenceNumber + ".msg", Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode);
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
    }
}
