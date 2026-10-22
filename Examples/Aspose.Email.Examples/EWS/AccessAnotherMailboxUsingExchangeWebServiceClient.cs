// Demonstrates working with a mailbox other than the one the client signed in to.
//
// GetMailboxInfo(mailbox) returns the folder URIs of another mailbox - a shared mailbox
// or a colleague's - provided the account has permission to open it; those URIs then
// work with every other call. Replace the address below with a mailbox of your
// organisation.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class AccessAnotherMailboxUsingExchangeWebServiceClient
    {
        public static void Run()
        {
            const string otherMailbox = "shared@example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var mailbox = client.GetMailboxInfo(otherMailbox);

                Console.WriteLine($"Mailbox of {otherMailbox}:");
                Console.WriteLine($"  Inbox:    {mailbox.InboxUri}");
                Console.WriteLine($"  Calendar: {mailbox.CalendarUri}");

                var inbox = client.GetFolderInfo(mailbox.InboxUri);
                Console.WriteLine($"  {inbox.TotalCount} message(s) in its Inbox, {inbox.UnreadCount} unread");
            }
        }
    }
}
