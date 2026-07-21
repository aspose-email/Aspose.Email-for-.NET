// Demonstrates how to convert an appointment EML to MSG while preserving the HTML body
// instead of converting it to RTF.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ConvertAppointmentEmlToMsgWithHtmlBody
    {
        public static void Run()
        {
            var mailMessage = MailMessage.Load(Data.Mapi/"TestAppointment.eml");

            var conversionOptions = new MapiConversionOptions
            {
                Format = OutlookMessageFormat.Unicode,
                ForcedRtfBodyForAppointment = false
            };

            var mapiMessage = MapiMessage.FromMailMessage(mailMessage, conversionOptions);
            Console.WriteLine("Body Type: " + mapiMessage.BodyType);

            mapiMessage.Save(Data.Out/"TestAppointment_out.msg");
        }
    }
}
