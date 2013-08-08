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

namespace MessagesFromIMAPServerToDisk
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

            client.SecurityMode = ImapSslSecurityMode.Implicit;

            // Enable SSL
            client.EnableSsl = true;

            try
            {
               
                client.Connect();
               
                //Log in to the remote server.
                client.Login();

                // Select the inbox folder
                client.SelectFolder(ImapFolderInfo.InBox);

                // Get the message info collection
                ImapMessageInfoCollection list = client.ListMessages();

                // Download each message
                for (int i = 0; i < list.Count; i++)
                {
                    //Save the EML file locally
                    client.SaveMessage(list[i].UniqueId, dataDir + list[i].UniqueId + ".eml");
                }

                //Disconnect to the remote IMAP server
                client.Disconnect();

                System.Console.WriteLine("Disconnected from the IMAP server");
            }
            catch (System.Exception ex)
            {
                System.Console.Write(ex.ToString());
            }
            
            
            
        }
    }
}