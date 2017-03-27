using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Calendar;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ManageAttachmentsFromCalendarFiles
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // ExStart:GetAttachmentsFromCalendar 
            string[] files = new string[3];
            files[0] = dataDir + "attachment_1.doc";
            files[1] = dataDir + "download.png";
            files[2] = dataDir + "Desert.jpg";

            Appointment app1 = new Appointment("Home", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1), "organizer@domain.com", "attendee@gmail.com");
            foreach (string file in files)
            {
                using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(file)))
                {
                    app1.Attachments.Add(new Attachment(ms, Path.GetFileName(file)));
                }
            }

            app1.Save(dataDir + "appWithAttachments_out.ics", AppointmentSaveFormat.Ics);

            Appointment app2 = Appointment.Load(dataDir + "appWithAttachments_out.ics");
            Console.WriteLine(app2.Attachments.Count);
            foreach (Attachment att in app2.Attachments)
                Console.WriteLine(att.Name);
            // ExEnd:GetAttachmentsFromCalendar
        }
    }
}