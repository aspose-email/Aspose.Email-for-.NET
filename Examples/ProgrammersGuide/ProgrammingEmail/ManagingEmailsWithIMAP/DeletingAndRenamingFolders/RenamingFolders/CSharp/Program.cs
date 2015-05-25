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

namespace RenamingFolders
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
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Rename a folder
                client.RenameFolder("Aspose", "Client");

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