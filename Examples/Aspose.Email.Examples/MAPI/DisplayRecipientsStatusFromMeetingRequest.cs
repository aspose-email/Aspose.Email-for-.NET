// Demonstrates how to load a meeting request MSG and display each recipient's tracking status.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class DisplayRecipientsStatusFromMeetingRequest
    {
        public static void Run()
        {
            var message = MapiMessage.Load(Data.Mapi/"Test Meeting.msg");

            foreach (var recipient in message.Recipients)
                Console.WriteLine(recipient.RecipientTrackStatus);
        }
    }
}
