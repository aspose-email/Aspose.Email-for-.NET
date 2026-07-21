// Demonstrates how to edit an existing message and save it back as a draft, so
// Outlook opens it for editing rather than showing it as sent.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SavingMessageInDraftStatus
    {
        public static void Run()
        {
            var msg = MailMessage.Load(Data.Mapi/"message.msg", new MsgLoadOptions());

            msg.Subject += "NEW SUBJECT (updated by Aspose.Email)";
            msg.HtmlBody += "NEW BODY (udpated by Aspose.Email)";

            var mapiMsg = MapiMessage.FromMailMessage(msg);

            // MSGFLAG_UNSENT is what marks the message as a draft.
            mapiMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT);

            var outputPath = Data.Out/"SavingMessageInDraftStatus_out.msg";
            mapiMsg.Save(outputPath);

            Console.WriteLine($"Subject: {mapiMsg.Subject}");
            Console.WriteLine("Saved as a draft (MSGFLAG_UNSENT), so Outlook opens it for editing.");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
