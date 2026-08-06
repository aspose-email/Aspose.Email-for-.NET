// Demonstrates how to tell an Outlook template (.oft) from an ordinary message
// (.msg). Both load into MapiMessage, and IsTemplate is what distinguishes them.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class DetectOutlookTemplate
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"message.msg");
            Console.WriteLine($"message.msg IsTemplate: {msg.IsTemplate}");

            var oft = MapiMessage.Load(Data.Mapi/"sample.oft");
            Console.WriteLine($"sample.oft  IsTemplate: {oft.IsTemplate}");

            // Saving with the OFT options turns an ordinary message into a template.
            var templatePath = Data.Out/"DetectOutlookTemplate_out.oft";
            msg.Save(templatePath, SaveOptions.DefaultOft);

            var saved = MapiMessage.Load(templatePath);
            Console.WriteLine($"Saved copy  IsTemplate: {saved.IsTemplate}");
            Console.WriteLine($"\nSaved to {templatePath}");
        }
    }
}
