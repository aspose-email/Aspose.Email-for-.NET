using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class RenamingFolders
    {
        public static void Run()
        {
            //ExStart:RenamingFolders
            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            // Specify host, username and password, port and SecurityOptions for your client
            client.Host = "imap.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 993;
            client.SecurityOptions = SecurityOptions.Auto;
            try
            {
                // Rename a folder and Disconnect to the remote IMAP server
                client.RenameFolder("Aspose", "Client");
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            //ExEnd:RenamingFolders
            Console.WriteLine(Environment.NewLine + "Renamed folders on IMAP server.");
        }
    }
}
