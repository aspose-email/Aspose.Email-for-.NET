// Demonstrates how the resource folder is referenced when a message is saved as HTML:
// relative paths keep the output portable, absolute ones tie it to this machine.
// The ResourceHtmlRendering event overrides the path entirely.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class SaveHtmlWithRelativeResourcePaths
    {
        public static void Run()
        {
            SaveWith("relative", true);
            SaveWith("absolute", false);
            SaveWithCustomPath();
        }

        private static void SaveWith(string label, bool useRelativePath)
        {
            var outputDir = Data.OutSub("Html-" + label);
            var msg = MailMessage.Load(Data.Email/"EmbeddedImage1.msg");

            var htmlSaveOptions = new HtmlSaveOptions
            {
                ResourceRenderingMode = ResourceRenderingMode.SaveToFile,
                UseRelativePathToResources = useRelativePath
            };

            var outputPath = outputDir/"target.html";
            msg.Save(outputPath, htmlSaveOptions);

            Console.WriteLine($"{label} paths -> {outputPath}");
            PrintFirstImageReference(outputPath);
        }

        private static void SaveWithCustomPath()
        {
            var outputDir = Data.OutSub("Html-custom");
            var msg = MailMessage.Load(Data.Email/"EmbeddedImage1.msg");

            var htmlSaveOptions = new HtmlSaveOptions
            {
                ResourceRenderingMode = ResourceRenderingMode.SaveToFile,
                UseRelativePathToResources = true
            };

            // The handler decides where each resource goes and how it is referenced.
            htmlSaveOptions.ResourceHtmlRendering += (o, args) =>
            {
                if (o is AttachmentBase attachment)
                    args.PathToResourceFile = Path.Combine("images", attachment.ContentType.Name);
            };

            var outputPath = outputDir/"target.html";
            msg.Save(outputPath, htmlSaveOptions);

            Console.WriteLine($"custom paths   -> {outputPath}");
            PrintFirstImageReference(outputPath);
        }

        private static void PrintFirstImageReference(string htmlPath)
        {
            foreach (var line in File.ReadLines(htmlPath))
            {
                var index = line.IndexOf("<img", StringComparison.OrdinalIgnoreCase);
                if (index < 0)
                    continue;

                var fragment = line.Substring(index);
                Console.WriteLine("  " + (fragment.Length <= 160 ? fragment : fragment.Substring(0, 160) + "..."));
                return;
            }
        }
    }
}
