// Demonstrates reading the basic facts about a mailbox.
//
// GetMailboxInfo returns the URIs of the well-known folders - Inbox, Sent Items,
// Drafts, Calendar, Contacts and so on - which every other call takes as input.
// GetMailboxSize returns the total size in bytes; GetFolderSizeRecursively shows how
// to measure single folders.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetMailboxInformationFromExchangeWebServices
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var mailbox = client.GetMailboxInfo();

                Console.WriteLine($"Mailbox size: {client.GetMailboxSize() / 1024.0 / 1024.0:F1} MB");
                Console.WriteLine($"Mailbox URI:  {mailbox.MailboxUri}");
                Console.WriteLine($"Inbox:        {mailbox.InboxUri}");
                Console.WriteLine($"Sent Items:   {mailbox.SentItemsUri}");
                Console.WriteLine($"Drafts:       {mailbox.DraftsUri}");
                Console.WriteLine($"Calendar:     {mailbox.CalendarUri}");
                Console.WriteLine($"Contacts:     {mailbox.ContactsUri}");
            }
        }
    }
}
