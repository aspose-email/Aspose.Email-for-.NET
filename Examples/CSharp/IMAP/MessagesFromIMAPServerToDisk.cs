using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.IMAP
{
    class MessagesFromIMAPServerToDisk
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_IMAP();
            string dstEmail = dataDir + "1234.eml";

            //Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            //Specify host, username and password for your client
            client.Host = "imap.gmail.com";

            // Set username
            client.Username = "your.username@gmail.com";

            // Set password
            client.Password = "your.password";

            // Set the port to 993. This is the SSL port of IMAP server
            client.Port = 993;

            // Enable SSL
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Select the inbox folder
                client.SelectFolder(ImapFolderInfo.InBox);

                // Get the message info collection
                ImapMessageInfoCollection list = client.ListMessages();

                // Download each message
                for (int i = 0; i < list.Count; i++)
                {
                    //Save the EML file locally
                    client.SaveMessage(list[i].UniqueId, dataDir + list[i].UniqueId + ".eml");
                }

                //Disconnect to the remote IMAP server
                client.Dispose();

            }
            catch (Exception ex)
            {
                System.Console.Write(Environment.NewLine + ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Downloaded messages from IMAP server.");
        }
    }
}
