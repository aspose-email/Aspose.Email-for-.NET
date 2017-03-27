using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class RecipientInformation
    {
        public static void Run()
        {
            // ExStart:RecipientInformation
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "Message.msg";

            // Load message file and Enumerate the recipients
            MapiMessage message = MapiMessage.FromFile(dstEmail);
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
            // ExEnd:RecipientInformation
            Console.WriteLine(Environment.NewLine + "Displayed recipient information from MSG file " + dstEmail);
        }
    }
}