using Aspose.Email.Calendar;
using Aspose.Email.Calendar.Recurrences;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class GeneratingOccurrencesFromRecurrencePatterns
    {        
        public static void Run()
        {
            GetOccurences();            
        }

        // ExStart:GeneratingOccurrencesFromRecurrencePatterns
        public static void GetOccurences()
        {
            // The path to the File directory
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();
            string tempFileName = dataDir + "Sample.pst";
            Appointment appointment = CreateAppointment();
            MailMessage mailMessage = CreateMessage();
            AlternateView alternateView = appointment.RequestApointment();
            mailMessage.AddAlternateView(alternateView);
            MapiMessage mapiMessage = MapiMessage.FromMailMessage(mailMessage);
            using (PersonalStorage pst = PersonalStorage.Create(tempFileName, FileFormatVersion.Unicode))
            {
                FolderInfo folder = pst.RootFolder.AddSubFolder("Calendar");
                folder.AddMessage(mapiMessage);
            }

            using (PersonalStorage pst = PersonalStorage.FromFile(tempFileName))
            {
                var folder = pst.RootFolder.GetSubFolder("Calendar");
                foreach (MessageInfo messageInfo in folder.GetContents())
                {
                    MapiMessage message = pst.ExtractMessage(messageInfo);
                    MapiCalendar meeting = (MapiCalendar)message.ToMapiMessageItem();
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        meeting.Save(memoryStream);
                        string s = StreamToString(memoryStream);
                        CalendarRecurrence recurrencePattern = new CalendarRecurrence(s);
                        DateCollection occurrences = recurrencePattern.GenerateOccurrences();
                        foreach (DateTime occurrence in occurrences)
                        {
                           Console.WriteLine("{0}", occurrence);
                        }
                    }
                }
            }

            File.Delete(tempFileName);
        }

        private static Appointment CreateAppointment()
        {
            MailAddressCollection attendees = new MailAddressCollection();
            attendees.Add(new MailAddress("attendee_address@aspose.com", "Attendee"));
            WeeklyRecurrencePattern expected = new WeeklyRecurrencePattern(3);
            Appointment app = new Appointment("Appointment Location", "Appointment Summary", "Appointment Description", DateTime.Now, DateTime.Now.AddMonths(1), "organizer_address@aspose.com", attendees, expected);
            
            return app;
        }

        private static MailMessage CreateMessage()
        {
            MailMessage mailMessage = new MailMessage("from_address@aspose.com", "to_address@aspose.com", "Mail Subject", "Mail Body");
            return mailMessage;
        }

        private static string StreamToString(Stream stream)
        {
            stream.Position = 0;
            StreamReader streamReader = new StreamReader(stream);
            return streamReader.ReadToEnd();
        }

        // ExEnd:GeneratingOccurrencesFromRecurrencePatterns
    }
}
