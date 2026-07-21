using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.IMAP
{
    class ImapBackupOperation
    {
        public static void Run()
        {
            // Create an instance of the ImapClient class
            ImapClient imapClient = new ImapClient();

            // Specify host, username and password, and set port for your client
            imapClient.Host = "imap.gmail.com";
            imapClient.Username = "your.username@gmail.com";
            imapClient.Password = "your.password";
            imapClient.Port = 993;
            imapClient.SecurityOptions = SecurityOptions.Auto;

            ImapMailboxInfo mailboxInfo = imapClient.MailboxInfo;

            ImapFolderInfo info = imapClient.GetFolderInfo(mailboxInfo.Inbox.Name);
            ImapFolderInfoCollection infos = new ImapFolderInfoCollection();
            infos.Add(info);

            imapClient.Backup(infos, Data.Out + @"\ImapBackup.pst", BackupOptions.Recursive);
        }
    }
}
