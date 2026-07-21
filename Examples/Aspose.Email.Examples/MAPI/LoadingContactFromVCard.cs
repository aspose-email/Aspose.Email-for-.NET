// Demonstrates how to load a contact from a VCard file using VCardContact and MapiContact.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.PersonalInfo.VCard;

namespace Aspose.Email.Examples.MAPI
{
    internal static class LoadingContactFromVCard
    {
        public static void Run()
        {
            // VCardContact exposes the file as-is, grouped by vCard property blocks.
            var vcfContact = VCardContact.Load(Data.Mapi/"Contact.vcf");
            Console.WriteLine($"VCard name: {vcfContact.IdentificationInfo.DisplayName}");

            // MapiContact maps the same file onto Outlook's contact model.
            var mapiContact = MapiContact.FromVCard(Data.Mapi/"Contact.vcf");
            Console.WriteLine($"MAPI name:  {mapiContact.NameInfo.DisplayName}");
        }
    }
}
