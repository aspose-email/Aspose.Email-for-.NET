// Demonstrates controlling whether letter case matters in a text search.
//
// The second argument of Contains is ignoreCase: true matches "newsletter",
// "Newsletter" and "NEWSLETTER" alike, false asks for the exact spelling given. Without
// the argument the server's default applies, which for IMAP is case-insensitive.
// Compare the counts to see the difference in your mailbox.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class CaseSensitiveEmailsFiltering
    {
        public static void Run()
        {
            const string text = "Newsletter";

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var builder = new ImapQueryBuilder();
                builder.Subject.Contains(text, true);
                var anyCase = client.ListMessages(builder.GetQuery());

                builder = new ImapQueryBuilder();
                builder.Subject.Contains(text, false);
                var exactCase = client.ListMessages(builder.GetQuery());

                Console.WriteLine($"Subject contains '{text}', any case:   {anyCase.Count} message(s)");
                Console.WriteLine($"Subject contains '{text}', exact case: {exactCase.Count} message(s)");

                foreach (var info in exactCase)
                    Console.WriteLine($"  {info.Subject}");
            }
        }
    }
}
