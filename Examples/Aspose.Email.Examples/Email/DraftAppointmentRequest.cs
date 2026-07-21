// Demonstrates how to create a draft appointment request by composing a MailMessage
// with an Appointment alternate view and saving it as an unsent MSG file.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.Email
{
    internal static class DraftAppointmentRequest
    {
        public static void Run()
        {
            var eml = new MailMessage("DanH@from.com", "KellyJ@to.com", string.Empty, string.Empty);

            var app = new Appointment(string.Empty, DateTime.Now, DateTime.Now, eml.From, eml.To)
            {
                MethodType = AppointmentMethodType.Publish
            };

            eml.AddAlternateView(app.RequestApointment());

            var msg = MapiMessage.FromMailMessage(eml);
            msg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT | MapiMessageFlags.MSGFLAG_FROMME);

            msg.Save(Data.Out/"appointment-draft_out.msg");
            Console.WriteLine($"Draft saved at {Data.Out/"appointment-draft_out.msg"}");
        }
    }
}
