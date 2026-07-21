// Demonstrates how to read voting buttons from a MapiMessage as a collection of strings.

using System;
using System.Collections;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadingOnlyVotingButtons
    {
        public static void Run()
        {
            MapiMessage testMsg = MapiMessage.Load(Data.Mapi/"MessageWithVotingButtons.msg");
            IList buttons = FollowUpManager.GetVotingButtons(testMsg);
            foreach (var button in buttons)
                Console.WriteLine("Button: " + button);
        }
    }
}
