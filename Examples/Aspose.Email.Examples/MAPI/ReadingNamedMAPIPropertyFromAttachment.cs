// Demonstrates how to read a custom named MAPI property from a message attachment.

using System;
using Aspose.Email;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadingNamedMapiPropertyFromAttachment
    {
        public static void Run()
        {
            string value = GetNamedPropertyByAspose();
            Console.WriteLine("CustomAttGuid: " + value);
        }

        private static string GetNamedPropertyByAspose()
        {
            MailMessage mail = MailMessage.Load(Data.Mapi/"outputAttachments.msg");
            var mapi = MapiMessage.FromMailMessage(mail);

            foreach (MapiNamedProperty namedProperty in mapi.Attachments[0].NamedProperties.Values)
            {
                if (string.Compare(namedProperty.NameId, "CustomAttGuid", StringComparison.OrdinalIgnoreCase) == 0)
                    return namedProperty.GetString();
            }
            return string.Empty;
        }
    }
}
