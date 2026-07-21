// Demonstrates how to create a MapiDistributionList and save it as MSG.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateAndSaveDistributionList
    {
        public static void Run()
        {
            var members = new MapiDistributionListMemberCollection();
            members.Add(new MapiDistributionListMember("John R. Patrick", "JohnRPatrick@armyspy.com"));
            members.Add(new MapiDistributionListMember("Tilly Bates", "TillyBates@armyspy.com"));

            var dlist = new MapiDistributionList("Simple list", members);
            dlist.Body = "Test body";
            dlist.Subject = "Test subject";
            dlist.Mileage = "Test mileage";
            dlist.Billing = "Test billing";

            var outputPath = Data.Out/"distlist_out.msg";
            dlist.Save(outputPath);

            Console.WriteLine($"Distribution list: {dlist.DisplayName}");
            foreach (var member in dlist.Members)
                Console.WriteLine($"    {member.DisplayName} <{member.EmailAddress}>");

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
