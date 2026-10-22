// Demonstrates filtering a folder on the server instead of downloading it.
//
// Build the conditions with ExchangeQueryBuilder, turn them into a MailQuery with
// GetQuery and pass it to ListMessages: Exchange evaluates the restriction and only the
// matching summaries come back. All conditions in one builder must hold at the same
// time.

using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class FilterMessagesUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var builder = new ExchangeQueryBuilder();
                builder.Subject.Contains("newsletter");
                builder.InternalDate.Since(DateTime.Today.AddDays(-30));

                var messages = client.ListMessages(client.MailboxInfo.InboxUri, builder.GetQuery(), false);

                Console.WriteLine($"'newsletter' in the subject, last 30 days: {messages.Count} message(s)");
                foreach (var info in messages)
                    Console.WriteLine($"  {info.InternalDate:d}  {info.From}  {info.Subject}");
            }
        }
    }
}
