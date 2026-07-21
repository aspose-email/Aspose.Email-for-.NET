// Demonstrates how to save an Outlook message as an Outlook template (OFT) file.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveMsgAsTemplate
    {
        public static void Run()
        {
            var templatePath = Data.Out/"SaveMsgAsTemplate_out.oft";

            using (var mapi = new MapiMessage("test@from.to", "test@to.to", "template subject", "Template body"))
            {
                // SaveOptions.DefaultOft selects the Outlook template format; saving without
                // options would produce an ordinary MSG file instead.
                mapi.Save(templatePath, SaveOptions.DefaultOft);
            }

            Console.WriteLine($"Saved Outlook template to {templatePath}");
        }
    }
}
