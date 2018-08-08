using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.IMAP
{
    public class AccessMailboxViaHTTPProxy
    {
        public static void Run()
        {
            //ExStart: AccessMailboxViaHTTPProxy
            HttpProxy proxy = new HttpProxy("18.222.124.59", 8080);
            using (ImapClient client = new ImapClient("imap.domain.com", "username", "password"))
            {
                client.Proxy = proxy;
                client.SelectFolder("Inbox");
            }
            //ExEnd: AccessMailboxViaHTTPProxy
        }
    }
}
