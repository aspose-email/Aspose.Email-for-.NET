// Demonstrates how to clear the follow-up flag from an Outlook message.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RemoveFollowUpflag
    {
        public static void Run()
        {
            var mapi = MapiMessage.Load(Data.Mapi/"message.msg");

            // The sample message carries no flag, so set one first - otherwise there
            // would be nothing for ClearFlag to remove.
            FollowUpManager.SetFlag(mapi, "Follow up");
            Console.WriteLine($"Flag before: {Describe(FollowUpManager.GetOptions(mapi).FlagRequest)}");

            FollowUpManager.ClearFlag(mapi);
            Console.WriteLine($"Flag after:  {Describe(FollowUpManager.GetOptions(mapi).FlagRequest)}");

            var outputPath = Data.Out/"RemoveFollowUpflag_out.msg";
            mapi.Save(outputPath);
            Console.WriteLine($"Saved to {outputPath}");
        }

        private static string Describe(string flagRequest) =>
            string.IsNullOrEmpty(flagRequest) ? "(not set)" : flagRequest;
    }
}
