using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class CreateREAndFWMessages
    {
        public static void Run()
        {
            var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission);

            try
            {
                var address = "user@domain";

                var eml = new MailMessage(address, address, "Please reply", "Please reply to me.");

                client.Send(eml);

                var builder = new ExchangeQueryBuilder();
                builder.Subject.Equals("Please reply", true);
                var query = builder.GetQuery();

                // Get list of messages
                var messageInfos = client.ListMessages(client.MailboxInfo.InboxUri, query, false);

                if (messageInfos != null && messageInfos.Count > 0)
                {
                    var reEml = new MailMessage(address, address, "Please reply1", "I replied!");

                    // Reply, Reply All and forward Message
                    client.Reply(reEml, messageInfos[0]);
                    client.ReplyAll(reEml, messageInfos[0]);
                    client.Forward(reEml, messageInfos[0]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in program"+ex.Message);
            }
        }
    }
}