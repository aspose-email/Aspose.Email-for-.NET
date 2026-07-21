using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.POP3
{
    class RecipientInformation
    {
        public static void Run()
        {
            string dstEmail = Data.Pop3 + "Message.msg";

            // Load message file and Enumerate the recipients
            MapiMessage message = MapiMessage.Load(dstEmail);
            foreach (MapiRecipient recip in message.Recipients)
            {
                switch (recip.RecipientType) // What's the type?
                {
                    case MapiRecipientType.MAPI_TO:
                        Console.WriteLine("RecipientType:TO");
                        break;
                    case MapiRecipientType.MAPI_CC:
                        Console.WriteLine("RecipientType:CC");
                        break;
                    case MapiRecipientType.MAPI_BCC:
                        Console.WriteLine("RecipientType:BCC");
                        break;
                }
                // Get email address, display name and  address type
                Console.WriteLine("Email Address: " + recip.EmailAddress);
                Console.WriteLine("DisplayName: " + recip.DisplayName);
                Console.WriteLine("AddressType: " + recip.AddressType);
            }
            Console.WriteLine(Environment.NewLine + "Displayed recipient information from MSG file " + dstEmail);
        }
    }
}