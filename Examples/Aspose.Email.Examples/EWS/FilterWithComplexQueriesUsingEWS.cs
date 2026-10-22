// Demonstrates combining search conditions with AND and OR.
//
// Every condition added to a builder must hold (AND). For alternatives, pass two
// conditions to Or; the result is a single condition that can be combined further, so
// nesting Or calls builds longer chains. Exchange evaluates the query on the server.

using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class FilterWithComplexQueriesUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var inboxUri = client.MailboxInfo.InboxUri;

                // AND: from example.com, during the last 7 days, but not today.
                var builder = new ExchangeQueryBuilder();
                builder.From.Contains("example.com");
                builder.InternalDate.Since(DateTime.Today.AddDays(-7));
                builder.InternalDate.Before(DateTime.Today);
                Console.WriteLine($"From example.com, last 7 days before today: " +
                                  $"{client.ListMessages(inboxUri, builder.GetQuery()).Count} message(s)");

                // OR: subject mentions "invoice" or the sender is billing@example.com.
                builder = new ExchangeQueryBuilder();
                builder.Or(builder.Subject.Contains("invoice"), builder.From.Contains("billing@example.com"));
                Console.WriteLine($"Invoice in subject or from billing:         " +
                                  $"{client.ListMessages(inboxUri, builder.GetQuery()).Count} message(s)");

                // Nested OR: any of three subjects.
                builder = new ExchangeQueryBuilder();
                builder.Or(
                    builder.Or(builder.Subject.Contains("order"), builder.Subject.Contains("shipment")),
                    builder.Subject.Contains("delivery"));
                Console.WriteLine($"Order, shipment or delivery in subject:     " +
                                  $"{client.ListMessages(inboxUri, builder.GetQuery()).Count} message(s)");
            }
        }
    }
}
