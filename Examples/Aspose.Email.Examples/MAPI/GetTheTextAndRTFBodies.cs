// Demonstrates how to retrieve and display the plain text and RTF body from a MapiMessage.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetTheTextAndRtfBodies
    {
        public static void Run()
        {
            var msg = MapiMessage.FromMailMessage(Data.Mapi/"Message.eml");

            if (msg.Body != null)
                Console.WriteLine(msg.Body);
            else
                Console.WriteLine("There's no text body.");

            if (msg.BodyRtf != null)
                Console.WriteLine(msg.BodyRtf);
            else
                Console.WriteLine("There's no RTF body.");
        }
    }
}
