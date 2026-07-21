// Demonstrates how to add follow-up options - start, due and reminder times - to
// an Outlook message.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetFollowUpflag
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

            var startDate = new DateTime(2013, 5, 23, 14, 40, 0);
            var reminderDate = new DateTime(2013, 5, 23, 16, 40, 0);
            var dueDate = reminderDate.AddDays(1);

            var options = new FollowUpOptions("Follow Up", startDate, dueDate, reminderDate);
            FollowUpManager.SetOptions(mapi, options);

            var outputPath = Data.Out/"SetFollowUpflag_out.msg";
            mapi.Save(outputPath);

            Console.WriteLine($"Flag:     {options.FlagRequest}");
            Console.WriteLine($"Start:    {startDate:g}");
            Console.WriteLine($"Due:      {dueDate:g}");
            Console.WriteLine($"Reminder: {reminderDate:g}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
