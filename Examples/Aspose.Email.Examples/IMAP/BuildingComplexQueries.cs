// Demonstrates combining search conditions with AND and OR.
//
// Every condition added to a builder must hold (AND). For alternatives, pass two
// conditions to Or; the result is a single condition that can be combined further, so
// nesting Or calls builds longer chains. The search runs on the server.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class BuildingComplexQueries
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                // AND: from example.com, received during the last 7 days, but not today.
                var builder = new ImapQueryBuilder();
                builder.From.Contains("example.com");
                builder.InternalDate.Since(DateTime.Today.AddDays(-7));
                builder.InternalDate.Before(DateTime.Today);

                var lastWeek = client.ListMessages(builder.GetQuery());
                Console.WriteLine($"From example.com, last 7 days before today: {lastWeek.Count} message(s)");

                // OR: subject mentions "invoice" or the sender is billing@example.com.
                builder = new ImapQueryBuilder();
                builder.Or(builder.Subject.Contains("invoice"), builder.From.Contains("billing@example.com"));

                var billing = client.ListMessages(builder.GetQuery());
                Console.WriteLine($"Invoice in subject or from billing:         {billing.Count} message(s)");

                // Nested OR: any of three subjects.
                builder = new ImapQueryBuilder();
                builder.Or(
                    builder.Or(builder.Subject.Contains("order"), builder.Subject.Contains("shipment")),
                    builder.Subject.Contains("delivery"));

                var orders = client.ListMessages(builder.GetQuery());
                Console.WriteLine($"Order, shipment or delivery in subject:     {orders.Count} message(s)");
            }
        }
    }
}
