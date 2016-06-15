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
    class AddingNewMessage
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_IMAP();
            string dstEmail = dataDir + "1234.eml";

            // Create a message
            Aspose.Email.Mail.MailMessage msg;
            msg = new Aspose.Email.Mail.MailMessage(
            "user@domain1.com",
            "user@domain2.com",
            "subject",
            "message"
            );

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
                // Subscribe to the Inbox folder
                client.SelectFolder(ImapFolderInfo.InBox);
                client.SubscribeFolder(client.CurrentFolder.Name);

                // Append the newly created message
                client.AppendMessage(client.CurrentFolder.Name, msg);

                System.Console.WriteLine("New Message Added Successfully");

                //Disconnect to the remote IMAP server
                client.Disconnect();

            }
            catch (Exception ex)
            {
                System.Console.Write(Environment.NewLine + ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Added new message on IMAP server.");
        }
    }
}
