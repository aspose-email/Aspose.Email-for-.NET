using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Exchange_EWS
{
    class SecondaryCalendarEvents
    {
        public static void Run()
        {
            // ExStart:SecondaryCalendarEvents
            using (IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "your.username", "your.Password"))
            {
                try
                {
                    // Create an appointmenta that will be added to secondary calendar folder

                    DateTime date = DateTime.Now;
                    DateTime startTime = new DateTime(date.Year, date.Month, date.Day, date.Hour, 0, 0);
                    DateTime endTime = startTime.AddHours(1);

                    string timeZone = "America/New_York";

                    Appointment[] listAppointments;
                    Appointment appointment = new Appointment("Room 121", startTime, endTime, "from@domain.com", "attendee@domain.com");
                    appointment.SetTimeZone(timeZone);
                    appointment.Summary = "EMAILNET-35198 - " + Guid.NewGuid().ToString();
                    appointment.Description = "EMAILNET-35198 Ability to add event to Secondary Calendar of Office 365";

                    // Verify that the new folder has been created
                    ExchangeFolderInfoCollection calendarSubFolders = client.ListSubFolders(client.MailboxInfo.CalendarUri);

                    string getfolderName;
                    string setFolderName = "New Calendar";
                    bool alreadyExits = false;

                    // Verify that the new folder has been created already exits or not 

                    for (int i = 0; i < calendarSubFolders.Count; i++)
                    {
                        getfolderName = calendarSubFolders[i].DisplayName;

                        if (getfolderName.Equals(setFolderName))
                        {
                            alreadyExits = true;
                        }
                    }

                    if (alreadyExits)
                    {
                        Console.WriteLine("Folder Already Exists");
                    }
                    else
                    {
                        // Create new calendar folder
                        client.CreateFolder(client.MailboxInfo.CalendarUri, setFolderName, null, "IPF.Appointment");

                        Console.WriteLine(calendarSubFolders.Count);

                        // Get the created folder URI
                        string newCalendarFolderUri = calendarSubFolders[0].Uri;

                        // appointment api with calendar folder uri
                        // Create
                        client.CreateAppointment(appointment, newCalendarFolderUri);
                        appointment.Location = "Room 122";
                        // update
                        client.UpdateAppointment(appointment, newCalendarFolderUri);
                        // list
                        listAppointments = client.ListAppointments(newCalendarFolderUri);

                        // list default calendar folder
                        listAppointments = client.ListAppointments(client.MailboxInfo.CalendarUri);

                        // Cancel
                        client.CancelAppointment(appointment, newCalendarFolderUri);
                        listAppointments = client.ListAppointments(newCalendarFolderUri);

                        // appointment api with context current calendar folder uri
                        client.CurrentCalendarFolderUri = newCalendarFolderUri;
                        // Create
                        client.CreateAppointment(appointment);
                        appointment.Location = "Room 122";
                        // update
                        client.UpdateAppointment(appointment);
                        // list
                        listAppointments = client.ListAppointments();

                        // list default calendar folder
                        listAppointments = client.ListAppointments(client.MailboxInfo.CalendarUri);

                        // Cancel
                        client.CancelAppointment(appointment);
                        listAppointments = client.ListAppointments();

                    }
                }
                finally
                {
                }
            }
            // ExEnd:SecondaryCalendarEvents
        }
    }
}
