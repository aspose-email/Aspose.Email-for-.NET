// Demonstrates unpacking an oledata.mso attachment.
//
// When Outlook sends a message whose body contains inline OLE content, it bundles that
// content into a single "oledata.mso" attachment rather than attaching each piece.
// InlineAttachmentExtractor opens that bundle and hands back the parts it holds, keyed
// by the name the body references them under.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ExtractInlineAttachmentsFromMsoPackage
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("MsoPackage");
            var found = 0;

            foreach (var path in new[] { Data.Mapi/"messageWithEmbeddedEML.msg", Data.Email/"Polymer homologue.msg" })
            {
                var msg = MapiMessage.Load(path);
                Console.WriteLine($"--- {Path.GetFileName(path)} ---");

                foreach (var attachment in msg.Attachments)
                {
                    var name = attachment.LongFileName ?? string.Empty;

                    if (!name.EndsWith(".mso", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"  {name}: ordinary attachment");
                        continue;
                    }

                    using (var stream = new MemoryStream(attachment.BinaryData))
                    {
                        var parts = InlineAttachmentExtractor.EnumerateMsoPackage(stream);
                        Console.WriteLine($"  {name}: {parts.Count} packaged part(s)");

                        foreach (var part in parts)
                        {
                            var outputPath = outputDir/$"{part.Key}.bin";
                            File.WriteAllBytes(outputPath, part.Value);

                            Console.WriteLine($"    {part.Key}: {part.Value.Length:N0} bytes -> {Path.GetFileName(outputPath)}");
                            found++;
                        }
                    }
                }
            }

            Console.WriteLine($"\nUnpacked {found} part(s) into {outputDir}");
        }
    }
}
