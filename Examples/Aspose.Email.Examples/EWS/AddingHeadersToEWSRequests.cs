using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Threading;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class AddingHeadersToEWSRequests
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.AddHeader("X-AnchorMailbox", "username@domain.com");
                var messageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri);
            }
        }
    }
}