using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ImapRestoreOperation
    {
        public static void Run()
        {
            //ExStart:1
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_IMAP();
            // Create an instance of the ImapClient class
            ImapClient imapClient = new ImapClient();

            // Specify host, username and password, and set port for your client
            imapClient.Host = "imap.gmail.com";
            imapClient.Username = "your.username@gmail.com";
            imapClient.Password = "your.password";
            imapClient.Port = 993;
            imapClient.SecurityOptions = SecurityOptions.Auto;

            RestoreSettings settings = new RestoreSettings();
            settings.Recursive = true;
            PersonalStorage pst = PersonalStorage.FromFile(dataDir + @"\ImapBackup.pst");
            imapClient.Restore(pst, settings);
            //ExEnd:1
        }
    }
}
