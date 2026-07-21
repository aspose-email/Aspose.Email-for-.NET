// Demonstrates how to create multiple MAPI contacts and save them to a PST contacts folder.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateNewMapiContactAndAddToContactsSubfolder
    {
        public static void Run()
        {
            // The constructors are shortcuts for the most common fields; anything beyond
            // them is set through the property sets, as contact4 to contact6 show.
            var contact1 = new MapiContact("Sebastian Wright", "SebastianWright@dayrep.com");
            var contact2 = new MapiContact("Wichert Kroos", "WichertKroos@teleworm.us", "Grade A Investment");
            var contact3 = new MapiContact("Christoffer van de Meeberg", "ChristoffervandeMeeberg@teleworm.us", "Krauses Sofa Factory", "046-630-4614046-630-4614");

            var contact4 = new MapiContact();
            contact4.NameInfo = new MapiContactNamePropertySet("Margaret", "J.", "Tolle");
            contact4.PersonalInfo.Gender = MapiContactGender.Female;
            contact4.ProfessionalInfo = new MapiContactProfessionalPropertySet("Adaptaz", "Recording engineer");
            contact4.PhysicalAddresses.WorkAddress.Address = "4 Darwinia Loop EIGHTY MILE BEACH WA 6725";
            contact4.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("Hisen1988", "SMTP", "MargaretJTolle@dayrep.com");
            contact4.Telephones.BusinessTelephoneNumber = "(08)9080-1183";
            contact4.Telephones.MobileTelephoneNumber = "(925)599-3355(925)599-3355";

            var contact5 = new MapiContact();
            contact5.NameInfo = new MapiContactNamePropertySet("Matthew", "R.", "Wilcox");
            contact5.PersonalInfo.Gender = MapiContactGender.Male;
            contact5.ProfessionalInfo = new MapiContactProfessionalPropertySet("Briazz", "Psychiatric aide");
            contact5.PhysicalAddresses.WorkAddress.Address = "Horner Strasse 12 4421 SAASS";
            contact5.Telephones.BusinessTelephoneNumber = "0650 675 73 300650 675 73 30";
            contact5.Telephones.HomeTelephoneNumber = "(661)387-5382(661)387-5382";

            var contact6 = new MapiContact();
            contact6.NameInfo = new MapiContactNamePropertySet("Bertha", "A.", "Buell");
            contact6.ProfessionalInfo = new MapiContactProfessionalPropertySet("Awthentikz", "Social work assistant");
            contact6.PersonalInfo.PersonalHomePage = "B2BTies.com";
            contact6.PhysicalAddresses.WorkAddress.Address = "Im Astenfeld 59 8580 EDELSCHROTT";
            contact6.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("Experwas", "SMTP", "BerthaABuell@armyspy.com");
            contact6.Telephones = new MapiContactTelephonePropertySet("06605045265");

            var path = Data.Out/"SampleContacts_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                var contactFolder = pst.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts);

                foreach (var contact in new[] { contact1, contact2, contact3, contact4, contact5, contact6 })
                {
                    contactFolder.AddMapiMessageItem(contact);
                    Console.WriteLine($"Added: {contact.NameInfo.DisplayName}");
                }

                Console.WriteLine($"\nAdded {contactFolder.ContentCount} contact(s) to {path}");
            }
        }
    }
}
