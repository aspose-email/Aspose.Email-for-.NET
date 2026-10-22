// Demonstrates controlling whether letter case matters in a text search.
//
// The second argument of Contains is ignoreCase: true matches "newsletter",
// "Newsletter" and "NEWSLETTER" alike, false asks for the exact spelling given.
// Compare the two counts to see the difference in your mailbox.

using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class CaseSensitiveEmailsFilteringUsingEWS
    {
        public static void Run()
        {
            const string text = "Newsletter";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var inboxUri = client.MailboxInfo.InboxUri;

                var builder = new ExchangeQueryBuilder();
                builder.Subject.Contains(text, true);
                var anyCase = client.ListMessages(inboxUri, builder.GetQuery(), false);

                builder = new ExchangeQueryBuilder();
                builder.Subject.Contains(text, false);
                var exactCase = client.ListMessages(inboxUri, builder.GetQuery(), false);

                Console.WriteLine($"Subject contains '{text}', any case:   {anyCase.Count} message(s)");
                Console.WriteLine($"Subject contains '{text}', exact case: {exactCase.Count} message(s)");
            }
        }
    }
}
