//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Outlook;
using System;

namespace GettingRecipientInformation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Load message file
            MapiMessage msg = MapiMessage.FromFile(dataDir + "Message.msg");

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
            
        }
    }
}