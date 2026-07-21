using Aspose.Email.Calendar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace Aspose.Email.Examples.EWS
{
    public class WriteMultipleEventsToICS
    {
        public static void Run()
        {
            AppointmentIcsSaveOptions saveOptions = new AppointmentIcsSaveOptions();
            saveOptions.Action = AppointmentAction.Create;
            using (CalendarWriter writer = new CalendarWriter(Data.Email + "WriteMultipleEventsToICS_out.ics", saveOptions))
            {
                for (int i = 0; i < 10; i++)
                {
                    Appointment app = new Appointment(string.Empty, DateTime.Now, DateTime.Now, "sender@domain.com", "receiver@domain.com");
                    app.Description = "Test body " + i;
                    app.Summary = "Test summary:" + i;
                    writer.Write(app);
                }
            }
        }
    }
}
