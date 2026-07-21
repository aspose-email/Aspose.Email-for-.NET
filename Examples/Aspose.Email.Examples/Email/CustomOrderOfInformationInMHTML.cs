// Demonstrates how to control the order of header fields rendered in MHTML output
// by customizing the RenderingHeaders collection.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class CustomOrderOfInformationInMhtml
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Attachments.eml");
            var opt = SaveOptions.DefaultMhtml;

            // An empty RenderingHeaders collection means the built-in default order.
            Save(eml, opt, "CustomOrderOfInformationInMHTML_1.mhtml", "default order");

            // Headers are rendered in the order they are added.
            opt.RenderingHeaders.Add(MhtTemplateName.From);
            opt.RenderingHeaders.Add(MhtTemplateName.Subject);
            opt.RenderingHeaders.Add(MhtTemplateName.To);
            opt.RenderingHeaders.Add(MhtTemplateName.Sent);
            Save(eml, opt, "CustomOrderOfInformationInMHTML_2.mhtml", "From, Subject, To, Sent");

            // Clear before building a different order, otherwise the two are combined.
            opt.RenderingHeaders.Clear();
            opt.RenderingHeaders.Add(MhtTemplateName.Attachments);
            opt.RenderingHeaders.Add(MhtTemplateName.Cc);
            opt.RenderingHeaders.Add(MhtTemplateName.Subject);
            Save(eml, opt, "CustomOrderOfInformationInMHTML_3.mhtml", "Attachments, Cc, Subject");
        }

        private static void Save(MailMessage eml, MhtSaveOptions opt, string fileName, string description)
        {
            var path = Data.Out/fileName;
            eml.Save(path, opt);
            Console.WriteLine($"{description,-24} -> {path}");
        }
    }
}
