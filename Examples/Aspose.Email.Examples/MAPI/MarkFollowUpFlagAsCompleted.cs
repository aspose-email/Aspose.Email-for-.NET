// Demonstrates how to mark the follow-up flag of a MapiMessage as completed.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class MarkFollowUpFlagAsCompleted
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"message.msg");

            // The sample message carries no flag, so set one first - otherwise there
            // would be nothing to complete.
            FollowUpManager.SetFlag(msg, "Follow up");
            Console.WriteLine($"Completed before: {FollowUpManager.GetOptions(msg).IsCompleted}");

            // Completing keeps the flag on the message and marks it done, rather than
            // removing it the way ClearFlag would.
            FollowUpManager.MarkAsCompleted(msg);
            Console.WriteLine($"Completed after:  {FollowUpManager.GetOptions(msg).IsCompleted}");

            var outputPath = Data.Out/"MarkedCompleted_out.msg";
            msg.Save(outputPath);
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
