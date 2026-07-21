// Demonstrates how to create a distribution list from PST contacts and from one-off members.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateDistributionListInPst
    {
        public static void Run()
        {
            // Distribution list from PST contacts
            var path1 = Data.Out/"CreateDistributionListInPST_out.pst";

            if (File.Exists(path1))
                File.Delete(path1);

            using (var pst = PersonalStorage.Create(path1, FileFormatVersion.Unicode))
            {
                var contactFolder = pst.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts);

                var entryId1 = contactFolder.AddMapiMessageItem(new MapiContact("Sebastian Wright", "SebastianWright@dayrep.com"));
                var entryId2 = contactFolder.AddMapiMessageItem(new MapiContact("Wichert Kroos", "WichertKroos@teleworm.us"));

                // Pointing the member at a contact's entry id links the two, so opening
                // the member in Outlook opens the underlying contact.
                var member1 = new MapiDistributionListMember("Sebastian Wright", "SebastianWright@dayrep.com")
                {
                    EntryIdType = MapiDistributionListEntryIdType.Contact,
                    EntryId = Convert.FromBase64String(entryId1)
                };

                var member2 = new MapiDistributionListMember("Wichert Kroos", "WichertKroos@teleworm.us")
                {
                    EntryIdType = MapiDistributionListEntryIdType.Contact,
                    EntryId = Convert.FromBase64String(entryId2)
                };

                var members = new MapiDistributionListMemberCollection { member1, member2 };

                var dlist = new MapiDistributionList("Contact list", members)
                {
                    Body = "Distribution List Body",
                    Subject = "Sample Distribution List using Aspose.Email"
                };

                contactFolder.AddMapiMessageItem(dlist);

                Console.WriteLine($"Linked to PST contacts: {dlist.DisplayName}, {members.Count} member(s)");
                Console.WriteLine($"  {path1}");
            }

            // Distribution list from one-off members (no separate contacts)
            var path2 = Data.Out/"CreateDistributionListInPST_OneOffmembers_out.pst";

            if (File.Exists(path2))
                File.Delete(path2);

            using (var pst = PersonalStorage.Create(path2, FileFormatVersion.Unicode))
            {
                var contactFolder = pst.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts);

                // One-off members carry their own name and address, so no contact has to
                // exist for them.
                var oneOffMembers = new MapiDistributionListMemberCollection
                {
                    new MapiDistributionListMember("John R. Patrick", "JohnRPatrick@armyspy.com"),
                    new MapiDistributionListMember("Tilly Bates", "TillyBates@armyspy.com")
                };

                var oneOffList = new MapiDistributionList("Simple list", oneOffMembers);
                contactFolder.AddMapiMessageItem(oneOffList);

                Console.WriteLine($"One-off members:        {oneOffList.DisplayName}, {oneOffMembers.Count} member(s)");
                Console.WriteLine($"  {path2}");
            }
        }
    }
}
