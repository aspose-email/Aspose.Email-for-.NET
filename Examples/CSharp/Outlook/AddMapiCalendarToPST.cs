using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System.IO;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class AddMapiCalendarToPST
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:AddMapiCalendarToPST
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create the appointment
            MapiCalendar appointment = new MapiCalendar(
                "LAKE ARGYLE WA 6743",
                "Appointment",
                "This is a very important meeting :)",
                new DateTime(2012, 10, 2, 13, 0, 0),
                new DateTime(2012, 10, 2, 14, 0, 0));

            // Create the meeting
            MapiRecipientCollection attendees = new MapiRecipientCollection();
            attendees.Add("ReneeAJones@armyspy.com", "Renee A. Jones", MapiRecipientType.MAPI_TO);
            attendees.Add("SzllsyLiza@dayrep.com", "Szollosy Liza", MapiRecipientType.MAPI_TO);

            MapiCalendar meeting = new MapiCalendar(
                "Meeting Room 3 at Office Headquarters",
                "Meeting",
                "Please confirm your availability.",
                new DateTime(2012, 10, 2, 13, 0, 0),
                new DateTime(2012, 10, 2, 14, 0, 0),
                "CharlieKhan@dayrep.com",
                attendees
                );

            string path = dataDir + "AddMapiCalendarToPST_out.pst";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (PersonalStorage pst = PersonalStorage.Create(dataDir + "AddMapiCalendarToPST_out.pst", FileFormatVersion.Unicode))
            {
                FolderInfo calendarFolder = pst.CreatePredefinedFolder("Calendar", StandardIpmFolder.Appointments);
                calendarFolder.AddMapiMessageItem(appointment);
                calendarFolder.AddMapiMessageItem(meeting);
            }
            // ExEnd:AddMapiCalendarToPST
        }
    }
}