using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.IMAP
{
    class DeletingAndRenamingFolders
    {
        public static void Run()
        {
            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            // Specify host, username and password, and set port for your client
            client.Host = "imap.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 993;
            client.SecurityOptions = SecurityOptions.Auto;
            try
            {
                // Delete a folder and Rename a folder
                client.DeleteFolder("foldername");
                client.RenameFolder("foldername", "newfoldername");

                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
        }
    }
}
