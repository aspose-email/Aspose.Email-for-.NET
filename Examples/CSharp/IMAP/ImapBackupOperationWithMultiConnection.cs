using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ImapBackupOperationWithMultiConnection
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

            imapClient.UseMultiConnection = MultiConnectionMode.Enable;

            ImapMailboxInfo mailboxInfo = imapClient.MailboxInfo;

            ImapFolderInfo info = imapClient.GetFolderInfo(mailboxInfo.Inbox.Name);
            ImapFolderInfoCollection infos = new ImapFolderInfoCollection();
            infos.Add(info);

            imapClient.Backup(infos, dataDir + @"\ImapBackup.pst", BackupOptions.Recursive);
            //ExEnd:1
        }
    }
}
