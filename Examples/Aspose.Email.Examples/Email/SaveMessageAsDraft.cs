// Demonstrates how to compose a MailMessage, convert it to a MapiMessage,
// and save it as an unsent draft MSG file.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class SaveMessageAsDraft
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                From = "from@domain.com",
                To = "to1@domain.com, to2@domain.com",
                Subject = "New message created by Aspose.Email",
                IsBodyHtml = true,
                HtmlBody = "<b>This line is in bold.</b> <br/> <br/><font color=blue>This line is in blue color</font>"
            };

            var msg = MapiMessage.FromMailMessage(eml);

            // MSGFLAG_UNSENT marks it as a draft, MSGFLAG_FROMME as sent by this user.
            msg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT | MapiMessageFlags.MSGFLAG_FROMME);

            var outputPath = Data.Out/"New-Draft.msg";
            msg.Save(outputPath);

            Console.WriteLine($"Draft: {msg.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
