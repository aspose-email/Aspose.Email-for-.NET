// Demonstrates how to flag a draft message for follow-up by its recipients, so the
// flag travels with the message rather than staying with the sender.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetFollowUpForRecipients
    {
        public static void Run()
        {
            var mailMsg = new MailMessage
            {
                Sender = "AETest12@gmail.com",
                To = "receiver@gmail.com",
                Body = "This message will test if follow up options can be added to a new mapi message."
            };

            var mapi = MapiMessage.FromMailMessage(mailMsg);

            // MSGFLAG_UNSENT marks the message as a draft; a recipient flag only makes
            // sense on a message that has not been sent yet.
            mapi.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT);

            var reminderDate = new DateTime(2013, 5, 23, 16, 40, 0);
            FollowUpManager.SetFlagForRecipients(mapi, "Follow up", reminderDate);

            var outputPath = Data.Out/"SetFollowUpForRecipients_out.msg";
            mapi.Save(outputPath);

            var options = FollowUpManager.GetOptions(mapi);
            Console.WriteLine($"Recipients flag: {options.RecipientsFlagRequest}");
            Console.WriteLine($"Reminder:        {reminderDate:g}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
