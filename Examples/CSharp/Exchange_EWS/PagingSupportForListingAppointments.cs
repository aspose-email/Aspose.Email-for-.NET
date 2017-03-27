using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.CSharp.Exchange_EWS
{
    class PagingSupportForListingAppointments
    {
        static void Run()
        { 
            // ExStart:PagingSupportForListingAppointments
            using (IEWSClient client = EWSClient.GetEWSClient("exchange.domain.com", "username", "password"))
            {
                try
                {
                    Appointment[] appts = client.ListAppointments();
                    Console.WriteLine(appts.Length);
                    DateTime date = DateTime.Now;
                    DateTime startTime = new DateTime(date.Year, date.Month, date.Day, date.Hour, 0, 0);
                    DateTime endTime = startTime.AddHours(1);
                    int appNumber = 10;
                    Dictionary<string, Appointment> appointmentsDict = new Dictionary<string, Appointment>();
                    for (int i = 0; i < appNumber; i++)
                    {
                        startTime = startTime.AddHours(1);
                        endTime = endTime.AddHours(1);
                        string timeZone = "America/New_York";
                        Appointment appointment = new Appointment(
                            "Room 112",
                            startTime,
                            endTime,
                            "from@domain.com",
                            "to@domain.com");
                        appointment.SetTimeZone(timeZone);
                        appointment.Summary = "NETWORKNET-35157_3 - " + Guid.NewGuid().ToString();
                        appointment.Description = "EMAILNET-35157 Move paging parameters to separate class";
                        string uid = client.CreateAppointment(appointment);
                        appointmentsDict.Add(uid, appointment);
                    }
                    AppointmentCollection totalAppointmentCol = client.ListAppointments();

                    ///// LISTING APPOINTMENTS WITH PAGING SUPPORT ///////
                    int itemsPerPage = 2;
                    List<AppointmentPageInfo> pages = new List<AppointmentPageInfo>();
                    AppointmentPageInfo pagedAppointmentCol = client.ListAppointmentsByPage(itemsPerPage);
                    Console.WriteLine(pagedAppointmentCol.Items.Count);
                    pages.Add(pagedAppointmentCol);
                    while (!pagedAppointmentCol.LastPage)
                    {
                        pagedAppointmentCol = client.ListAppointmentsByPage(itemsPerPage, pagedAppointmentCol.PageOffset + 1);
                        pages.Add(pagedAppointmentCol);
                    }
                    int retrievedItems = 0;
                    foreach (AppointmentPageInfo folderCol in pages)
                        retrievedItems += folderCol.Items.Count;
                }
                finally
                {
                }
            }
            // ExEnd:PagingSupportForListingAppointments
        }
    }
}
