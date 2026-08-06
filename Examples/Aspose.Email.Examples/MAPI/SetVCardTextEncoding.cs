// Demonstrates how to control the text encoding a vCard is written in, which is what
// keeps non-ASCII names readable in the resulting file.

using System;
using System.Text;
using Aspose.Email.PersonalInfo.VCard;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetVCardTextEncoding
    {
        public static void Run()
        {
            var contact = VCardContact.Load(Data.Mapi/"Contact.vcf");
            Console.WriteLine($"Loaded: {contact.IdentificationInfo.DisplayName}");

            var options = new VCardSaveOptions
            {
                PreferredTextEncoding = Encoding.UTF8
            };

            var outputPath = Data.Out/"SetVCardTextEncoding_out.vcf";
            contact.Save(outputPath, options);

            Console.WriteLine($"Saved with {options.PreferredTextEncoding.WebName} to {outputPath}");
        }
    }
}
