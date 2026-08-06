// Demonstrates how to turn a multi-contact VCF file into an Outlook distribution
// list, so a contact export can be used as a mailing list directly.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateDistributionListFromVCF
    {
        public static void Run()
        {
            // Build a multi-contact VCF by concatenating the sample contact twice.
            var single = File.ReadAllText(Data.Mapi/"Contact.vcf");
            var multiPath = Data.Out/"CreateDistributionListFromVCF_in.vcf";
            File.WriteAllText(multiPath, single + Environment.NewLine + single);

            var dlist = MapiDistributionList.FromVCF(multiPath);

            Console.WriteLine($"Distribution list: {dlist.Subject}");
            Console.WriteLine($"Members: {dlist.Members.Count}");
            foreach (var member in dlist.Members)
                Console.WriteLine($"  {member.DisplayName} <{member.EmailAddress}>");

            var outputPath = Data.Out/"CreateDistributionListFromVCF_out.msg";
            dlist.Save(outputPath);
            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
