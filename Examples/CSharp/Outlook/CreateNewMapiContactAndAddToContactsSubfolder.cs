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
    class CreateNewMapiContactAndAddToContactsSubfolder
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:CreateNewMapiContactAndAddToContactsSubfolder
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create three Contacts 
            MapiContact contact1 = new MapiContact("Sebastian Wright", "SebastianWright@dayrep.com");
            MapiContact contact2 = new MapiContact("Wichert Kroos", "WichertKroos@teleworm.us", "Grade A Investment");
            MapiContact contact3 = new MapiContact("Christoffer van de Meeberg", "ChristoffervandeMeeberg@teleworm.us", "Krauses Sofa Factory", "046-630-4614046-630-4614");

            // Contact #4
            MapiContact contact4 = new MapiContact();
            contact4.NameInfo = new MapiContactNamePropertySet("Margaret", "J.", "Tolle");
            contact4.PersonalInfo.Gender = MapiContactGender.Female;
            contact4.ProfessionalInfo = new MapiContactProfessionalPropertySet("Adaptaz", "Recording engineer");
            contact4.PhysicalAddresses.WorkAddress.Address = "4 Darwinia Loop EIGHTY MILE BEACH WA 6725";
            contact4.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("Hisen1988", "SMTP", "MargaretJTolle@dayrep.com");
            contact4.Telephones.BusinessTelephoneNumber = "(08)9080-1183";
            contact4.Telephones.MobileTelephoneNumber = "(925)599-3355(925)599-3355";

            // Contact #5
            MapiContact contact5 = new MapiContact();
            contact5.NameInfo = new MapiContactNamePropertySet("Matthew", "R.", "Wilcox");
            contact5.PersonalInfo.Gender = MapiContactGender.Male;
            contact5.ProfessionalInfo = new MapiContactProfessionalPropertySet("Briazz", "Psychiatric aide");
            contact5.PhysicalAddresses.WorkAddress.Address = "Horner Strasse 12 4421 SAASS";
            contact5.Telephones.BusinessTelephoneNumber = "0650 675 73 300650 675 73 30";
            contact5.Telephones.HomeTelephoneNumber = "(661)387-5382(661)387-5382";

            // Contact #6
            MapiContact contact6 = new MapiContact();
            contact6.NameInfo = new MapiContactNamePropertySet("Bertha", "A.", "Buell");
            contact6.ProfessionalInfo = new MapiContactProfessionalPropertySet("Awthentikz", "Social work assistant");
            contact6.PersonalInfo.PersonalHomePage = "B2BTies.com";
            contact6.PhysicalAddresses.WorkAddress.Address = "Im Astenfeld 59 8580 EDELSCHROTT";
            contact6.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("Experwas", "SMTP", "BerthaABuell@armyspy.com");
            contact6.Telephones = new MapiContactTelephonePropertySet("06605045265");

            // Load the Outlook file
            string path = dataDir + "SampleContacts_out.pst";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (PersonalStorage personalStorage = PersonalStorage.Create(dataDir + "SampleContacts_out.pst", FileFormatVersion.Unicode))
            {
                FolderInfo contactFolder = personalStorage.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts);
                contactFolder.AddMapiMessageItem(contact1);
                contactFolder.AddMapiMessageItem(contact2);
                contactFolder.AddMapiMessageItem(contact3);
                contactFolder.AddMapiMessageItem(contact4);
                contactFolder.AddMapiMessageItem(contact5);
                contactFolder.AddMapiMessageItem(contact6);
            }
            // ExEnd:CreateNewMapiContactAndAddToContactsSubfolder
        }
    }
}