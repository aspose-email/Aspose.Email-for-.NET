// Demonstrates how to rename an inline image of an HTML message. The name lives in
// the resource's ContentDisposition, separate from the content id the body links to.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ChangeLinkedResourceFileName
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"EmbeddedImage1.msg");
            Console.WriteLine($"Linked resources: {eml.LinkedResources.Count}");

            for (var i = 0; i < eml.LinkedResources.Count; i++)
            {
                var resource = eml.LinkedResources[i];
                Console.WriteLine($"Before: {resource.ContentDisposition.FileName} (content id: {resource.ContentId})");

                resource.ContentDisposition.FileName = $"changed{i}.png";
                Console.WriteLine($"After:  {resource.ContentDisposition.FileName}");
            }

            var outputPath = Data.Out/"ChangeLinkedResourceFileName_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);
            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
