//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using System;
using Aspose.Email.Imap;

namespace GettingFoldersInformation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

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
                //Connect to the remote server.
                client.Connect();               

                //Log in to the remote server.
                client.Login();

                // Get all folders in the currently subscribed folder
                Aspose.Email.Imap.ImapFolderInfoCollection folderInfoColl = client.ListFolders();

                // Iterate through the collection to get folder info one by one
                foreach (Aspose.Email.Imap.ImapFolderInfo folderInfo in folderInfoColl)
                {
                    // Folder name
                    Console.WriteLine("Folder name is " + folderInfo.Name);
                    ImapFolderInfo folderExtInfo = client.ListFolder(folderInfo.Name);
                    // New messages in the folder
                    Console.WriteLine("New message count: " + folderExtInfo.NewMessageCount);
                    // Check whether its readonly
                    Console.WriteLine("Is it readonly? " + folderExtInfo.ReadOnly);
                    // Total number of messages
                    Console.WriteLine("Total number of messages " + folderExtInfo.TotalMessageCount);
                }

                //Disconnect to the remote IMAP server
                client.Disconnect();
                
            }
            catch (Exception ex)
            {
                System.Console.Write(ex.ToString());
            }
        }
    }
}