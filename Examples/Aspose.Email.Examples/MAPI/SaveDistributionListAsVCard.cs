// Demonstrates how to export an Outlook distribution list as a single multi-contact
// VCF file, one vCard per member.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.PersonalInfo.VCard;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveDistributionListAsVCard
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"NewGroup.msg");
            var dlist = (MapiDistributionList)msg.ToMapiMessageItem();

            Console.WriteLine($"Distribution list: {dlist.Subject}");
            Console.WriteLine($"Members: {dlist.Members.Count}");

            // ContactSaveFormat.VCard writes every member into one file.
            var options = new MapiDistributionListSaveOptions(ContactSaveFormat.VCard);

            var outputPath = Data.Out/"SaveDistributionListAsVCard_out.vcf";
            dlist.Save(outputPath, options);

            Console.WriteLine($"Written as multi-contact VCF: {VCardContact.IsMultiContacts(outputPath)}");
            Console.WriteLine($"File size: {new FileInfo(outputPath).Length} byte(s)");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
