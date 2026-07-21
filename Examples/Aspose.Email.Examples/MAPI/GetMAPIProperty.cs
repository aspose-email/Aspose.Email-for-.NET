// Demonstrates how to access specific MAPI properties (PR_SUBJECT and PR_INTERNET_CPID) from an MSG file.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetMapiProperty
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"message.msg");

            // Try ANSI subject first, fall back to Unicode peer
            var prop = msg.Properties[MapiPropertyTag.PR_SUBJECT]
                    ?? msg.Properties[MapiPropertyTag.PR_SUBJECT_W];

            if (prop == null)
            {
                Console.WriteLine("No property found!");
                return;
            }

            Console.WriteLine("Subject: " + prop.GetString());

            var codePage = msg.Properties[MapiPropertyTag.PR_INTERNET_CPID];
            if (codePage != null)
                Console.WriteLine("CodePage: " + codePage.GetLong());
        }
    }
}
