using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class DeletingFolders
    {
        public static void Run()
        {
            //ExStart:DeletingFolders
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
                // Rename a folder and Disconnect to the remote IMAP server
                client.DeleteFolder("Client");
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            //ExEnd:DeletingFolders
            Console.WriteLine(Environment.NewLine + "Deleted folders on IMAP server.");
        }
    }
}
