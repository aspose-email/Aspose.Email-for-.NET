//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;

namespace CSharp.POP3
{
    class RecipientInformation
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "message.msg";

            //Load message file
            MapiMessage msg = MapiMessage.FromFile(dstEmail);

            //Enumerate the recipients
            foreach (MapiRecipient recip in msg.Recipients)
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

                //Get email address
                Console.WriteLine("Email Address: " + recip.EmailAddress);
                //Get display name
                Console.WriteLine("DisplayName: " + recip.DisplayName);
                //Get address type
                Console.WriteLine("AddressType: " + recip.AddressType);
            }

            Console.WriteLine(Environment.NewLine + "Displayed recipient information from MSG file " + dstEmail);
        }
    }
}
