// Demonstrates searching with Advanced Query Syntax (AQS), the language of the Outlook
// search box.
//
// AQS searches use the mailbox's search index, which makes them fast on large folders
// and lets them look into attachments. ExchangeAdvancedSyntaxQueryBuilder builds the
// query in code; ExchangeAdvancedSyntaxMailQuery takes a ready-made AQS string such as
// "subject:invoice hasattachment:true".

using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class FilterMessagesOnCriteriaWithAqsUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var inboxUri = client.MailboxInfo.InboxUri;

                var builder = new ExchangeAdvancedSyntaxQueryBuilder();
                builder.Subject.Contains("invoice");
                builder.HasAttachment.Equals(true);

                var built = client.ListMessages(inboxUri, builder.GetQuery());
                Console.WriteLine($"Built with the builder:  {built.Count} message(s)");
                foreach (var info in built)
                    Console.WriteLine($"  {info.Subject}");

                var query = new ExchangeAdvancedSyntaxMailQuery("subject:invoice hasattachment:true");
                var fromString = client.ListMessages(inboxUri, query);
                Console.WriteLine($"Written as an AQS string: {fromString.Count} message(s)");
            }
        }
    }
}
