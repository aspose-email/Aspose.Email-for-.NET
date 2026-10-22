// Demonstrates common search criteria: arrival date, sender, recipient, Message-ID and
// size.
//
// Each query uses a fresh builder, and all conditions added to one builder must hold.
// MailQueryBuilder holds the criteria every protocol understands; ExchangeQueryBuilder
// derives from it and adds Exchange-specific ones such as MessageId and ItemSize.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.EWS
{
    internal static class FilterMessagesOnCriteriaUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var builder = new MailQueryBuilder();
                builder.InternalDate.On(DateTime.Now);
                Search(client, "Arrived today", builder.GetQuery());

                builder = new MailQueryBuilder();
                builder.InternalDate.Since(DateTime.Now.AddDays(-7));
                Search(client, "Arrived in the last 7 days", builder.GetQuery());

                builder = new MailQueryBuilder();
                builder.From.Contains("example.com");
                Search(client, "From anyone at example.com", builder.GetQuery());

                builder = new MailQueryBuilder();
                builder.To.Contains("recipient@example.com");
                Search(client, "Sent to recipient@example.com", builder.GetQuery());

                var exchangeBuilder = new ExchangeQueryBuilder();
                exchangeBuilder.MessageId.Equals("<message-id@example.com>");
                Search(client, "With a given Message-ID", exchangeBuilder.GetQuery());

                exchangeBuilder = new ExchangeQueryBuilder();
                exchangeBuilder.ItemSize.Greater(1024 * 1024);
                Search(client, "Larger than 1 MB", exchangeBuilder.GetQuery());
            }
        }

        private static void Search(IEWSClient client, string description, MailQuery query)
        {
            var messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
            Console.WriteLine($"{description,-32} {messages.Count} message(s)");
        }
    }
}
