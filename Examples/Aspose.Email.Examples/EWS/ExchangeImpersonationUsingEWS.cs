// Demonstrates Exchange impersonation: acting as another user.
//
// An account with the ApplicationImpersonation role - typically a service account - can
// call ImpersonateUser to make every following request run as the given user, as if
// that user had signed in. ResetImpersonation switches back. Replace the address below
// with a user of your organisation; the example only reads.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ExchangeImpersonationUsingEWS
    {
        public static void Run()
        {
            const string otherUser = "colleague@example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                Report(client, "Own mailbox:");

                client.ImpersonateUser(ItemChoice.PrimarySmtpAddress, otherUser);
                Report(client, $"Impersonating {otherUser}:");

                client.ResetImpersonation();
                Report(client, "Back to the own mailbox:");
            }
        }

        private static void Report(IEWSClient client, string title)
        {
            var inbox = client.GetFolderInfo(client.GetMailboxInfo().InboxUri);
            Console.WriteLine($"{title} {inbox.TotalCount} message(s) in the Inbox");
        }
    }
}
