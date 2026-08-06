// Demonstrates how to handle a VCF file that holds more than one contact:
// IsMultiContacts tells the two cases apart, LoadAsMultiple reads them all.

using System;
using System.IO;
using Aspose.Email.PersonalInfo.VCard;

namespace Aspose.Email.Examples.MAPI
{
    internal static class LoadMultipleContactsFromVCard
    {
        public static void Run()
        {
            // Build a multi-contact VCF by concatenating the sample contact twice, so the
            // example has something to read regardless of the shipped sample data.
            var single = File.ReadAllText(Data.Mapi/"Contact.vcf");
            var multiPath = Data.Out/"LoadMultipleContactsFromVCard_out.vcf";
            File.WriteAllText(multiPath, single + Environment.NewLine + single);

            Console.WriteLine($"Contact.vcf holds multiple contacts: {VCardContact.IsMultiContacts(Data.Mapi/"Contact.vcf")}");
            Console.WriteLine($"The generated file holds multiple contacts: {VCardContact.IsMultiContacts(multiPath)}");

            if (VCardContact.IsMultiContacts(multiPath))
            {
                var contacts = VCardContact.LoadAsMultiple(multiPath, new VCardLoadOptions());

                Console.WriteLine($"\nLoaded {contacts.Count} contact(s):");
                foreach (var contact in contacts)
                    Console.WriteLine($"  {contact.IdentificationInfo.DisplayName}");
            }
        }
    }
}
