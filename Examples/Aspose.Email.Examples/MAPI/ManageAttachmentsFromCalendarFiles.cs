// Demonstrates how to add file attachments to an Appointment and read them back from the saved ICS.

using System;
using System.IO;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ManageAttachmentsFromCalendarFiles
    {
        public static void Run()
        {
            var app = new Appointment("Home", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1),
                "organizer@domain.com", "attendee@gmail.com");

            foreach (var file in new[] { "attachment_1.doc", "download.png", "Desert.jpg" })
                app.Attachments.Add(new Attachment(new MemoryStream(File.ReadAllBytes(Data.Mapi/file)), file));

            app.Save(Data.Out/"appWithAttachments_out.ics", AppointmentSaveFormat.Ics);

            var app2 = Appointment.Load(Data.Out/"appWithAttachments_out.ics");
            Console.WriteLine("Attachment count: " + app2.Attachments.Count);

            foreach (var att in app2.Attachments)
                Console.WriteLine(att.Name);
        }
    }
}
