// Demonstrates how to read the PR_SUBJECT and PR_INTERNET_CPID MAPI properties from a MapiMessage.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadiNamedMapiPropertyFromAttachment
    {
        public static void Run()
        {
            MapiMessage msg = MapiMessage.Load(Data.Mapi/"message.msg");

            MapiProperty prop = msg.Properties[MapiPropertyTag.PR_SUBJECT]
                ?? msg.Properties[MapiPropertyTag.PR_SUBJECT_W];

            if (prop == null)
            {
                Console.WriteLine("No property found!");
                return;
            }

            Console.WriteLine("Subject:" + prop.GetString());

            MapiProperty codePage = msg.Properties[MapiPropertyTag.PR_INTERNET_CPID];
            if (codePage != null)
                Console.WriteLine("CodePage:" + codePage.GetLong());
        }
    }
}
