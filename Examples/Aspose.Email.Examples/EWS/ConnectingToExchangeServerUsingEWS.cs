// Demonstrates connecting to Exchange Online over EWS.
//
// ClientBuilder.Ews reads the "Ews" section of clientsettings.json, obtains an OAuth
// token for the registered application and passes it to EWSClient.GetEWSClient. The
// client loads the mailbox's well-known folder URIs on connect; MailboxInfo exposes
// them for all further calls.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class ConnectingToExchangeServerUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var mailbox = client.MailboxInfo;

                Console.WriteLine($"Connected to {mailbox.MailboxUri}");
                Console.WriteLine($"Exchange version: {client.ServerVersion}");

                var inbox = client.GetFolderInfo(mailbox.InboxUri);
                Console.WriteLine($"{inbox.DisplayName}: {inbox.TotalCount} item(s), {inbox.UnreadCount} unread");
            }
        }
    }
}
