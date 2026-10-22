// Demonstrates reading entries of the organisation's global address list.
//
// ListMailboxes returns the mailboxes the directory knows, optionally narrowed by a
// filter; ResolveContact turns an ambiguous name into the matching entries, like the
// "Check Names" button in Outlook. Both return MapiContact objects. Directory look-ups
// over EWS return a limited number of entries per call (GetMailboxes, for instance, at
// most 100), so narrow them down with a filter in large organisations.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class ListGlobalAddressList
    {
        public static void Run()
        {
            const string ambiguousName = "john";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var mailboxes = client.ListMailboxes();
                Console.WriteLine($"{mailboxes.Count} mailbox(es) listed:");
                foreach (var mailbox in mailboxes)
                    Console.WriteLine("  " + Describe(mailbox));

                var matches = client.ResolveContact(ambiguousName);
                Console.WriteLine($"\n'{ambiguousName}' resolves to {matches.Count} match(es):");
                foreach (var match in matches)
                    Console.WriteLine("  " + Describe(match));
            }
        }

        private static string Describe(MapiContact contact)
        {
            return $"{contact.NameInfo?.DisplayName} <{contact.ElectronicAddresses?.Email1?.EmailAddress}>";
        }
    }
}
