// Demonstrates how to read the reactions (thumbs up, heart, ...) that recipients
// left on an Outlook message.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetMessageReactions
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"message.msg");

            var reactions = FollowUpManager.GetReactions(msg);
            Console.WriteLine($"Reactions on '{msg.Subject}': {reactions.Count}");

            foreach (var reaction in reactions)
            {
                Console.WriteLine($"  {reaction.Name} <{reaction.Email}> " +
                                  $"reacted with {reaction.Type} on {reaction.ReactionDateTime}");
            }

            if (reactions.Count == 0)
                Console.WriteLine("  (this sample message carries none - the API returns an empty list)");
        }
    }
}
