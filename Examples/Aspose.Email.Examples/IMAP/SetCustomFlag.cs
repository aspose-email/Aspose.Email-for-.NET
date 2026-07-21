using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.IMAP
{
    class SetCustomFlag
    {
        public static void Run()
        {
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
                // Create a message
                MailMessage message = new MailMessage("user@domain1.com", "user@domain2.com", "subject", "message");

                //Append the message to mailbox
                string uid = client.AppendMessage(ImapFolderInfo.InBox, message);

                //Add custom flags to the added messge
                client.AddMessageFlags(uid, ImapMessageFlags.Keyword("custom1") | ImapMessageFlags.Keyword("custom1_0"));

                //Retreive the messages for checking the presence of custom flag
                client.SelectFolder(ImapFolderInfo.InBox);

                ImapMessageInfoCollection messageInfos = client.ListMessages();
                foreach (var inf in messageInfos)
                {
                    ImapMessageFlags[] flags = inf.Flags.Split();

                    if (inf.ContainsKeyword("custom1"))
                        Console.WriteLine("Keyword found");
                }

                Console.WriteLine("Setting Custom Flag to Message example executed successfully!");
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
        }
    }
}
