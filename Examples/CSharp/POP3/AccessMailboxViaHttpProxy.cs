using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.POP3
{
    class AccessMailboxViaHttpProxy
    {
        public static void Run()
        {
            //ExStart: AccessMailboxViaHTTPProxy
            HttpProxy proxy = new HttpProxy("18.222.124.59", 8080);
            using (Pop3Client client = new Pop3Client("imap.domain.com", "username", "password"))
            {
                client.Proxy = proxy;
                Pop3MailboxInfo mailboxInfo = client.GetMailboxInfo();
            }
            //ExEnd: AccessMailboxViaHTTPProxy
        }
    }
}
