using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SettingMessageFlags
    {
        public static void Run()
        {            
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
                
                // ExStart:SettingMessageFlags
                // Mark the message as read
                client.ChangeMessageFlags(1, ImapMessageFlags.IsRead);
                // ExEnd:SettingMessageFlags

                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }            
            Console.WriteLine(Environment.NewLine + "Set message flags from IMAP server.");
        }
    }
}
