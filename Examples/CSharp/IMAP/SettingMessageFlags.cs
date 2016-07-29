using System;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SettingMessageFlags
    {
        public static void Run()
        {
            //ExStart:SettingMessageFlags
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
                client.ChangeMessageFlags(1, ImapMessageFlags.IsRead);
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            //ExEnd:SettingMessageFlags
            Console.WriteLine(Environment.NewLine + "Set message flags from IMAP server.");
        }
    }
}
