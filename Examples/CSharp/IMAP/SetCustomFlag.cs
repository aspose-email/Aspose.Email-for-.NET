using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
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
                //ExStart:SetCustomFlag
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

                //ExEnd:SetCustomFlag

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
