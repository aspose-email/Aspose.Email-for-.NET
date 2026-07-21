using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class ExchangeFoldersBackupToPST
    {
        public static void Run()
        {
            // Create instance of IEWSClient class by providing credentials
            const string mailboxUri = "https://ews.domain.com/ews/Exchange.asmx";
            const string domain = @"";
            const string username = @"username";
            const string password = @"password";
            NetworkCredential credential = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credential);

            // Get Exchange mailbox info of other email account
            ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();
            ExchangeFolderInfo info = client.GetFolderInfo(mailboxInfo.InboxUri);
            ExchangeFolderInfoCollection fc = new ExchangeFolderInfoCollection();
            fc.Add(info);
            client.Backup(fc, Data.Out + "Backup_out.pst", BackupOptions.None);
        }
    }
}