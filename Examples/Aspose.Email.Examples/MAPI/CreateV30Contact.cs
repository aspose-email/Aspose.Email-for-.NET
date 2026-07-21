// Demonstrates how to create a MAPI contact and save it as a vCard 3.0 VCF file.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.PersonalInfo.VCard;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateV30Contact
    {
        public static void Run()
        {
            var contact = new MapiContact();
            contact.NameInfo = new MapiContactNamePropertySet("Jane", "A.", "Buell");
            contact.ProfessionalInfo = new MapiContactProfessionalPropertySet("Aspose Pty Ltd", "Social work assistant");
            contact.PersonalInfo.PersonalHomePage = "Aspose.com";
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("test@test.com");
            contact.Telephones.HomeTelephoneNumber = "06605040000";

            // Without the options the contact would be written as vCard 2.1, which is
            // what Outlook produces by default.
            var saveOptions = new VCardSaveOptions { Version = VCardVersion.V30 };

            var outputPath = Data.Out/"V30.vcf";
            contact.Save(outputPath, saveOptions);

            Console.WriteLine($"Contact: {contact.NameInfo.DisplayName}");
            Console.WriteLine($"vCard version: {saveOptions.Version}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
