using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System.IO;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreateDistributionListInPST
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:CreateDistributionListInPST
            string dataDir = RunExamples.GetDataDir_Outlook();

            string displayName1 = "Sebastian Wright";
            string email1 = "SebastianWright@dayrep.com";

            string displayName2 = "Wichert Kroos";
            string email2 = "WichertKroos@teleworm.us";

            string strEntryId1;
            string strEntryId2;

            string path = dataDir + "CreateDistributionListInPST_out.pst";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            // Create distribution list from contacts
            using (PersonalStorage personalStorage = PersonalStorage.Create(dataDir + "CreateDistributionListInPST_out.pst", FileFormatVersion.Unicode))
            {
                // Add the contact folder to pst
                FolderInfo contactFolder = personalStorage.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts);

                // Create contacts
                strEntryId1 = contactFolder.AddMapiMessageItem(new MapiContact(displayName1, email1));
                strEntryId2 = contactFolder.AddMapiMessageItem(new MapiContact(displayName2, email2));

                // Create distribution list on the base of the created contacts
                MapiDistributionListMember member1 = new MapiDistributionListMember(displayName1, email1);
                member1.EntryIdType = MapiDistributionListEntryIdType.Contact;
                member1.EntryId = Convert.FromBase64String(strEntryId1);

                MapiDistributionListMember member2 = new MapiDistributionListMember(displayName2, email2);
                member2.EntryIdType = MapiDistributionListEntryIdType.Contact;
                member2.EntryId = Convert.FromBase64String(strEntryId2);

                MapiDistributionListMemberCollection members = new MapiDistributionListMemberCollection();
                members.Add(member1);
                members.Add(member2);

                MapiDistributionList distributionList = new MapiDistributionList("Contact list", members);
                distributionList.Body = "Distribution List Body";
                distributionList.Subject = "Sample Distribution List using Aspose.Email";

                // Add distribution list to PST
                contactFolder.AddMapiMessageItem(distributionList);
            }


            string path1 = dataDir + "CreateDistributionListInPST_OneOffmembers_out.pst";

            if (File.Exists(path1))
            {
                File.Delete(path1);
            }

            // Create one-off distribution list members (for which no separate contacts were created)
            using (PersonalStorage personalStorage = PersonalStorage.Create(dataDir + "CreateDistributionListInPST_OneOffmembers_out.pst", FileFormatVersion.Unicode))
            {
                // Add the contact folder to pst
                FolderInfo contactFolder = personalStorage.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts);

                MapiDistributionListMemberCollection oneOffmembers = new MapiDistributionListMemberCollection();
                oneOffmembers.Add(new MapiDistributionListMember("John R. Patrick", "JohnRPatrick@armyspy.com"));
                oneOffmembers.Add(new MapiDistributionListMember("Tilly Bates", "TillyBates@armyspy.com"));

                MapiDistributionList oneOffMembersList = new MapiDistributionList("Simple list", oneOffmembers);
                contactFolder.AddMapiMessageItem(oneOffMembersList);
            }
            // ExEnd:CreateDistributionListInPST
        }
    }
}