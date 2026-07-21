// Demonstrates how to set MAPI properties on a message and on one of its
// recipients, including a date/time property.

using System;
using System.Text;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetMAPIProperties
    {
        public static void Run()
        {
            var mapiMsg = new MapiMessage("user1@gmail.com", "user2@gmail.com", "This is subject", "This is body");

            // String properties are stored as raw bytes, so the encoding has to match the
            // property tag: the _W suffix means Unicode.
            mapiMsg.SetProperty(new MapiProperty(
                MapiPropertyTag.PR_SENDER_ADDRTYPE_W, Encoding.Unicode.GetBytes("EX")));

            // Properties can also be set per recipient - here the recipient is addressed
            // as a fax rather than by email.
            var recipientTo = mapiMsg.Recipients[0];
            recipientTo.SetProperty(new MapiProperty(
                MapiPropertyTag.PR_RECEIVED_BY_ADDRTYPE_W, Encoding.Unicode.GetBytes("MYFAX")));

            const string faxAddress = "My Fax User@/FN=fax#/VN=voice#/CO=My Company/CI=Local";
            recipientTo.SetProperty(new MapiProperty(
                MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS_W, Encoding.Unicode.GetBytes(faxAddress)));

            // Flags are a bit field: mark the message as an unsent message from this user.
            mapiMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT | MapiMessageFlags.MSGFLAG_FROMME);
            mapiMsg.SetProperty(new MapiProperty(MapiPropertyTag.PR_RTF_IN_SYNC, BitConverter.GetBytes((long)1)));

            // CreateMapiPropertyFromDateTime handles the conversion to the 8-byte
            // FILETIME layout that PT_SYSTIME properties expect.
            var modificationTime = new DateTime(2013, 9, 11);
            mapiMsg.SetProperty(MapiProperty.CreateMapiPropertyFromDateTime(
                MapiPropertyTag.PR_LAST_MODIFICATION_TIME, modificationTime));

            var outputPath = Data.Out/"MapiProp_out.msg";
            mapiMsg.Save(outputPath);

            Console.WriteLine($"Sender address type: {mapiMsg.SenderAddressType}");
            Console.WriteLine($"Last modified:       {modificationTime:d}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
