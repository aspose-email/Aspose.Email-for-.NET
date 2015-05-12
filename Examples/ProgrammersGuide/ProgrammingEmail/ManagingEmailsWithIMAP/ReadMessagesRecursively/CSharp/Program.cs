//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Imap;
using System;
using Aspose.Email.Mail;

namespace ReadMessagesRecursively
{
    public class Program
    {
        public static void Main(string[] args)
        {

            //Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            //Specify host, username and password for your client
            client.Host = "imap.gmail.com";

            // Set username
            client.Username = "asposetest123@gmail.com";

            // Set password
            client.Password = "F123456f";

            // Set the port to 993. This is the SSL port of IMAP server
            client.Port = 993;

            // Enable SSL
            client.EnableSsl = true;

            try
            {
                System.Console.WriteLine("Connecting to the IMAP server");

                //Connect to the remote server.
                client.Connect();

                System.Console.WriteLine("Connected to the IMAP server");

                //Log in to the remote server.
                client.Login();

                System.Console.WriteLine("Logged in to the IMAP server");

                // The root folder (which will be created on disk) consists of host and username
                string rootFolder = client.Host + "-" + client.Username;
                // Create the root folder
                Directory.CreateDirectory(rootFolder);

                // List all the folders from IMAP server
                ImapFolderInfoCollection folderInfoCollection = client.ListFolders();
                foreach (ImapFolderInfo folderInfo in folderInfoCollection)
                {
                    // Call the recursive method to read messages and get sub-folders
                    ListMessagesInFolder(folderInfo, rootFolder,client);
                }


                //Disconnect to the remote IMAP server
                client.Disconnect();

                System.Console.WriteLine("Disconnected from the IMAP server");
            }
            catch (Exception ex)
            {
                System.Console.Write(ex.ToString());
            }
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
                ImapFolderInfo folderInfoStatus = client.ListFolder(folderInfo.Name);
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
                    msg.Save(currentFolder + "\\" + fileName + "-" + msgInfo.SequenceNumber + ".msg", MailMessageSaveType.OutlookMessageFormat);
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
                    ListMessagesInFolder(subfolderInfo, rootFolder,client);
                }
            }
            catch (Exception) { }
        }
 
    }
}