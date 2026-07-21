// Demonstrates how to create a detailed MAPI contact and save it as both MSG and VCard.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateAndSaveOutlookContact
    {
        public static void Run()
        {
            var contact = new MapiContact();
            contact.NameInfo = new MapiContactNamePropertySet("Bertha", "A.", "Buell");
            contact.ProfessionalInfo = new MapiContactProfessionalPropertySet("Awthentikz", "Social work assistant");
            contact.PersonalInfo.PersonalHomePage = "B2BTies.com";
            contact.PhysicalAddresses.WorkAddress.Address = "Im Astenfeld 59 8580 EDELSCHROTT";
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("Experwas", "SMTP", "BerthaABuell@armyspy.com");
            contact.Telephones = new MapiContactTelephonePropertySet("06605045265");
            contact.PersonalInfo.Children = new string[] { "child1", "child2", "child3" };
            contact.Categories = new string[] { "category1", "category2", "category3" };
            contact.Mileage = "Some test mileage";
            contact.Billing = "Test billing information";
            contact.OtherFields.Journal = true;
            contact.OtherFields.Private = true;
            contact.OtherFields.ReminderTime = new DateTime(2014, 1, 1, 0, 0, 55);
            contact.OtherFields.ReminderTopic = "Test topic";
            contact.OtherFields.UserField1 = "ContactUserField1";
            contact.OtherFields.UserField2 = "ContactUserField2";
            contact.OtherFields.UserField3 = "ContactUserField3";
            contact.OtherFields.UserField4 = "ContactUserField4";

            contact.Photo = new MapiContactPhoto(
                File.ReadAllBytes(Data.Mapi/"Desert.jpg"),
                MapiContactPhotoImageFormat.Jpeg);

            // MSG keeps every Outlook-specific field; VCard keeps only what the vCard
            // standard defines, so some of the fields above are dropped from the .vcf.
            var msgPath = Data.Out/"MapiContact_out.msg";
            var vcfPath = Data.Out/"MapiContact_out.vcf";
            contact.Save(msgPath, ContactSaveFormat.Msg);
            contact.Save(vcfPath, ContactSaveFormat.VCard);

            Console.WriteLine($"Contact: {contact.NameInfo.DisplayName}");
            Console.WriteLine($"Company: {contact.ProfessionalInfo.CompanyName}, {contact.ProfessionalInfo.Title}");
            Console.WriteLine($"Email:   {contact.ElectronicAddresses.Email1.EmailAddress}");
            Console.WriteLine($"Saved to {msgPath}");
            Console.WriteLine($"Saved to {vcfPath}");
        }
    }
}
