// Demonstrates searching for text that is not plain ASCII.
//
// IMAP sends search strings with a charset. Passing an Encoding to the ImapQueryBuilder
// constructor makes the client send the conditions in that charset - UTF-8 covers every
// language - so the server can match words with accented or non-Latin letters.

using System;
using System.Text;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SpecifyEncodingForQueryBuilder
    {
        public static void Run()
        {
            // A German and a Russian greeting, written with escapes to keep the file ASCII.
            var words = new[] { "Gr\u00fc\u00dfe", "\u041f\u0440\u0438\u0432\u0435\u0442" };

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                foreach (var word in words)
                {
                    var builder = new ImapQueryBuilder(Encoding.UTF8);
                    builder.Subject.Contains(word);

                    var messages = client.ListMessages(builder.GetQuery());
                    Console.WriteLine($"Subject contains '{word}': {messages.Count} message(s)");
                    foreach (var info in messages)
                        Console.WriteLine($"  {info.Subject}");
                }
            }
        }
    }
}
