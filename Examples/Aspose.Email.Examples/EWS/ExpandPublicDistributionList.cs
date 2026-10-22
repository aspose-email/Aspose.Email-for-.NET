// Demonstrates listing the members of an organisation-wide distribution group.
//
// ExpandDistributionList takes the e-mail address of a distribution group from the
// global address list and returns its members, as the "+" next to a group name in
// Outlook does. Replace the address below with a group of your organisation.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class ExpandPublicDistributionList
    {
        public static void Run()
        {
            const string groupAddress = "all-staff@example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var members = client.ExpandDistributionList(new MailAddress(groupAddress));

                Console.WriteLine($"{groupAddress} has {members.Count} member(s):");
                foreach (var member in members)
                    Console.WriteLine($"  {member.DisplayName} <{member.Address}>");
            }
        }
    }
}
