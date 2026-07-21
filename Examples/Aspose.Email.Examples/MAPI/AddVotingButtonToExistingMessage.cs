// Demonstrates how to add a voting button to an existing MSG message.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddVotingButtonToExistingMessage
    {
        public static void Run()
        {
            var mapi = MapiMessage.Load(Data.Mapi/"message.msg");
            Console.WriteLine($"Buttons before: {Describe(mapi)}");

            FollowUpManager.AddVotingButton(mapi, "Indeed!");
            Console.WriteLine($"Buttons after:  {Describe(mapi)}");

            var outputPath = Data.Out/"AddVotingButtonToExistingMessage_out.msg";
            mapi.Save(outputPath);
            Console.WriteLine($"Saved to {outputPath}");
        }

        private static string Describe(MapiMessage msg)
        {
            var buttons = FollowUpManager.GetVotingButtons(msg);
            return buttons.Length == 0 ? "(none)" : string.Join(", ", buttons);
        }
    }
}
