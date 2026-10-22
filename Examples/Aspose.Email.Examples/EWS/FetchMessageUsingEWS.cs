// Demonstrates downloading a complete message.
//
// FetchMessage takes the UniqueUri from a listing and returns the full message as a
// MailMessage - body, headers and attachments - ready to read, convert or save.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class FetchMessageUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (var info in client.ListMessages(client.MailboxInfo.InboxUri, 5))
                {
                    var message = client.FetchMessage(info.UniqueUri);

                    Console.WriteLine($"{message.Subject}");
                    Console.WriteLine($"  from {message.From}, {message.Attachments.Count} attachment(s)");
                    foreach (var attachment in message.Attachments)
                        Console.WriteLine($"    {attachment.Name}");
                }
            }
        }
    }
}
