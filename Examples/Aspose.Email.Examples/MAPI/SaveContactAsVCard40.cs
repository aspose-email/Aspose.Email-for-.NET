// Demonstrates how to write a MAPI contact as vCard 4.0. The version matters because
// fields such as children and hobbies are only carried by the newer formats.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.PersonalInfo.VCard;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveContactAsVCard40
    {
        public static void Run()
        {
            var contact = new MapiContact
            {
                NameInfo = new MapiContactNamePropertySet("Jane", "A.", "Buell")
            };
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("jane.buell@example.com");
            contact.PersonalInfo.Children = new[] { "child 1", "child 2" };
            contact.PersonalInfo.Hobbies = "Hobbies";
            contact.PersonalInfo.Notes = "Notes";

            foreach (var version in new[] { VCardVersion.V30, VCardVersion.V40 })
            {
                var outputPath = Data.Out/$"SaveContactAsVCard_{version}.vcf";
                contact.Save(outputPath, new VCardSaveOptions { Version = version });

                Console.WriteLine($"{version} -> {outputPath}");
            }
        }
    }
}
