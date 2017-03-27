using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SetMAPIProperties
    {
        public static void Run()
        {
            //ExStart:SetMAPIProperties
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create a sample Message
            MapiMessage mapiMsg = new MapiMessage("user1@gmail.com", "user2@gmail.com", "This is subject", "This is body");

            // Set multiple properties
            mapiMsg.SetProperty(new MapiProperty(MapiPropertyTag.PR_SENDER_ADDRTYPE_W, Encoding.Unicode.GetBytes("EX")));
            MapiRecipient recipientTo = mapiMsg.Recipients[0];
            MapiProperty propAddressType = new MapiProperty(MapiPropertyTag.PR_RECEIVED_BY_ADDRTYPE_W, Encoding.UTF8.GetBytes("MYFAX"));
            recipientTo.SetProperty(propAddressType);
            string faxAddress = "My Fax User@/FN=fax#/VN=voice#/CO=My Company/CI=Local";
            MapiProperty propEmailAddress = new MapiProperty(MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS_W, Encoding.UTF8.GetBytes(faxAddress));
            recipientTo.SetProperty(propEmailAddress);
            mapiMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT | MapiMessageFlags.MSGFLAG_FROMME);
            mapiMsg.SetProperty(new MapiProperty(MapiPropertyTag.PR_RTF_IN_SYNC, BitConverter.GetBytes((long)1)));

            // Set DateTime property
            MapiProperty modificationTime = new MapiProperty(MapiPropertyTag.PR_LAST_MODIFICATION_TIME, ConvertDateTime(new DateTime(2013, 9, 11)));
            mapiMsg.SetProperty(modificationTime);
            mapiMsg.Save(dataDir + "MapiProp_out.msg");
            //ExEnd:SetMAPIProperties
        }

        //ExStart:ConvertDateTime
        private static byte[] ConvertDateTime(DateTime t)
        {
            long filetime = t.ToFileTime();
            byte[] d = new byte[8];
            d[0] = (byte)(filetime & 0xFF);
            d[1] = (byte)((filetime & 0xFF00) >> 8);
            d[2] = (byte)((filetime & 0xFF0000) >> 16);
            d[3] = (byte)((filetime & 0xFF000000) >> 24);
            d[4] = (byte)((filetime & 0xFF00000000) >> 32);
            d[5] = (byte)((filetime & 0xFF0000000000) >> 40);
            d[6] = (byte)((filetime & 0xFF000000000000) >> 48);
            d[7] = (byte)(((ulong)filetime & 0xFF00000000000000) >> 56);
            return d;
        }
        //ExEnd:ConvertDateTime
    }
}