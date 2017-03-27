using System;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
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
    class CreateAndSaveCalendaritems
    {
        public static void Run()
        {
            //ExStart:CreateAndSaveCalendaritems
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create the appointment
            MapiCalendar calendar = new MapiCalendar(
                "LAKE ARGYLE WA 6743",
                "Appointment",
                "This is a very important meeting :)",
                new DateTime(2012, 10, 2, 13, 0, 0),
                new DateTime(2012, 10, 2, 14, 0, 0));

            calendar.Save(dataDir + "CalendarItem_out.ics", AppointmentSaveFormat.Ics);
            //ExEnd:CreateAndSaveCalendaritems
        }
    }
}