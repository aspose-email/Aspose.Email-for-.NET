// Demonstrates that the producing application's identifier survives a conversion:
// the PRODID of a vCard or an ICS file is carried into MapiContact.ProductId and
// MapiCalendar.ProductId, and either can be overridden when saving.

using System;
using System.IO;
using Aspose.Email.Calendar;
using Aspose.Email.Mapi;
using Aspose.Email.PersonalInfo.VCard;

namespace Aspose.Email.Examples.MAPI
{
    internal static class PreserveProductIdOnConversion
    {
        public static void Run()
        {
            RoundTripContact();
            Console.WriteLine();
            RoundTripCalendar();
        }

        private static void RoundTripContact()
        {
            var contact = VCardContact.Load(Data.Mapi/"Contact.vcf");

            // The sample vCard carries no PRODID, so set one to have something to follow
            // through the conversion. A vCard exported by Outlook would already have it.
            contact.ExplanatoryInfo.ProdId = "-//Aspose Ltd//Aspose.Email Examples//EN";

            var prodId = contact.ExplanatoryInfo.ProdId;
            Console.WriteLine($"vCard PRODID: {Show(prodId)}");

            // Round-trip the contact through MSG and read the identifier back.
            using (var ms = new MemoryStream())
            {
                contact.Save(ms, ContactSaveFormat.Msg);
                ms.Position = 0;

                var mapiContact = (MapiContact)MapiMessage.Load(ms).ToMapiMessageItem();
                Console.WriteLine($"MapiContact.ProductId: {Show(mapiContact.ProductId)}");
                Console.WriteLine($"Preserved: {prodId == mapiContact.ProductId}");

                // The save options override it for the next hop.
                var options = new MapiContactSaveOptions { ProductId = "New ProductId" };
                var outputPath = Data.Out/"PreserveProductIdOnConversion_contact.msg";
                mapiContact.Save(outputPath, options);
                Console.WriteLine($"Saved with ProductId '{options.ProductId}' to {outputPath}");
            }
        }

        private static void RoundTripCalendar()
        {
            var app = Appointment.Load(Data.Email/"test.ics");
            var prodId = app.ProductId;
            Console.WriteLine($"ICS PRODID: {Show(prodId)}");

            using (var ms = new MemoryStream())
            {
                app.Save(ms, new AppointmentMsgSaveOptions());
                ms.Position = 0;

                var mapiCalendar = (MapiCalendar)MapiMessage.Load(ms).ToMapiMessageItem();
                Console.WriteLine($"MapiCalendar.ProductId: {Show(mapiCalendar.ProductId)}");
                Console.WriteLine($"Preserved: {prodId == mapiCalendar.ProductId}");

                // The calendar options spell it ProductIdentifier.
                var options = new MapiCalendarMsgSaveOptions { ProductIdentifier = "New ProductId" };
                var outputPath = Data.Out/"PreserveProductIdOnConversion_calendar.msg";
                mapiCalendar.Save(outputPath, options);
                Console.WriteLine($"Saved with ProductIdentifier '{options.ProductIdentifier}' to {outputPath}");
            }
        }

        private static string Show(string value) => string.IsNullOrEmpty(value) ? "(none)" : value;
    }
}
