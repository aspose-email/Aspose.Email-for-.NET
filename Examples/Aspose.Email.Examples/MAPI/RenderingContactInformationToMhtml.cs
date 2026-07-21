// Demonstrates how to render a vCard contact to MHTML, choosing which groups of
// contact fields end up in the output.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RenderingContactInformationToMhtml
    {
        public static void Run()
        {
            var contact = MapiContact.FromVCard(Data.Mapi/"Contact.vcf");

            // MHTML is produced by MailMessage, so the contact goes through MSG on its
            // way to a message: contact -> MSG -> MailMessage -> MHTML.
            using (var ms = new MemoryStream())
            {
                contact.Save(ms, ContactSaveFormat.Msg);
                ms.Position = 0;

                var eml = MapiMessage.Load(ms).ToMailMessage(new MailConversionOptions());

                var mhtSaveOptions = new MhtSaveOptions
                {
                    CheckBodyContentEncoding = true,
                    PreserveOriginalBoundaries = true,

                    // RenderVCardInfo is what turns the contact fields into the body;
                    // without it only the message header would be written.
                    MhtFormatOptions = MhtFormatOptions.WriteHeader | MhtFormatOptions.RenderVCardInfo,

                    // Only these groups of fields are rendered.
                    RenderedContactFields = ContactFieldsSet.NameInfo | ContactFieldsSet.PersonalInfo |
                                            ContactFieldsSet.Telephones | ContactFieldsSet.Events
                };

                var outputPath = Data.Out/"ContactMhtml_out.mhtml";
                eml.Save(outputPath, mhtSaveOptions);

                Console.WriteLine($"Rendered contact: {contact.NameInfo.DisplayName}");
                Console.WriteLine($"Saved to {outputPath}");
            }
        }
    }
}
