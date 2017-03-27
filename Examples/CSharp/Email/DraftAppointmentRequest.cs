using System.IO;
using System;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.CSharp.Email
{
    class DraftAppointmentRequest
    {
        public static void Run()
        {
            // ExStart:DraftAppointmentRequest
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstDraft = dataDir + "appointment-draft_out.msg";
            
            string sender = "test@gmail.com";
            string recipient = "test@email.com";

            MailMessage message = new MailMessage(sender, recipient, string.Empty, string.Empty);

            Appointment app = new Appointment(string.Empty, DateTime.Now, DateTime.Now, sender, recipient);
            app.Method = AppointmentMethodType.Publish;

            message.AddAlternateView(app.RequestApointment());

            MapiMessage msg = MapiMessage.FromMailMessage(message);

            // Save the appointment as draft.
            msg.Save(dstDraft);

            Console.WriteLine(Environment.NewLine + "Draft saved at " + dstDraft);
            // ExEnd:DraftAppointmentRequest
        }
    }
}
