using System;
using Aspose.Email.Imap;
using Aspose.Email.Mail;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class AddingNewMessage
    {
        public static void Run()
        {
            //ExStart:AddingNewMessage
            // Create a message
            MailMessage msg = new MailMessage("user@domain1.com", "user@domain2.com", "subject", "message");

            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            // Specify host, username, password, port and SecurityOptions for your client
            client.Host = "imap.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 993;
            client.SecurityOptions = SecurityOptions.Auto;
            try
            {
                // Subscribe to the Inbox folder, Append the newly created message and Disconnect to the remote IMAP server
                client.SelectFolder(ImapFolderInfo.InBox);
                client.SubscribeFolder(client.CurrentFolder.Name);
                client.AppendMessage(client.CurrentFolder.Name, msg);
                Console.WriteLine("New Message Added Successfully");
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            Console.WriteLine(Environment.NewLine + "Added new message on IMAP server.");
            //ExEnd:AddingNewMessage
        }
    }
}
