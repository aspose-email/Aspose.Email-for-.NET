// Demonstrates reading the private distribution lists of a mailbox and their members.
//
// ListDistributionLists returns the contact groups of the Contacts folder;
// FetchDistributionList returns the member addresses of one of them.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class FetchPrivateDistributionList
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var lists = client.ListDistributionLists();
                Console.WriteLine($"{lists.Length} private distribution list(s):");

                foreach (var list in lists)
                {
                    var members = client.FetchDistributionList(list);
                    Console.WriteLine($"  {list.DisplayName} ({members.Count} member(s))");
                    foreach (var member in members)
                        Console.WriteLine($"    {member.Address}");
                }
            }
        }
    }
}
