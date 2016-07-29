using System;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class RemovingMessageFlags
    {
        public static void Run()
        {
            // ExStart:RemovingMessageFlags
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
                Console.WriteLine("Logged in to the IMAP server");
                // Mark the message as read and Disconnect to the remote IMAP server
                client.RemoveMessageFlags(1, ImapMessageFlags.IsRead);
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            // ExEnd:RemovingMessageFlags
            Console.WriteLine(Environment.NewLine + "Removed message flags from IMAP server.");
        }
    }
}
