using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.IMAP
{
    public class ImapSpecialUseMailboxes
    {
        public static void Run()
        {
            //ExStart: 1
            ImapClient imapClient = new ImapClient();
            imapClient.Host = "<HOST>";
            imapClient.Port = 993;
            imapClient.Username = "<USERNAME>";
            imapClient.Password = "<PASSWORD>";
            imapClient.SupportedEncryption = EncryptionProtocols.Tls;
            imapClient.SecurityOptions = SecurityOptions.SSLImplicit;
            
            ImapMailboxInfo mailboxInfo = imapClient.MailboxInfo;
            Console.WriteLine(mailboxInfo.Inbox);
            Console.WriteLine(mailboxInfo.DraftMessages);
            Console.WriteLine(mailboxInfo.JunkMessages);
            Console.WriteLine(mailboxInfo.SentMessages);
            Console.WriteLine(mailboxInfo.Trash);
            //ExEnd: 1

            Console.WriteLine("ImapSpecialUseMailboxes executed successfully.");
        }
    }
}
