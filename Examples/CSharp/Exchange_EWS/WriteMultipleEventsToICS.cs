using Aspose.Email.Calendar;
using Aspose.Email.Examples.CSharp.Email;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Exchange_EWS
{
    public class WriteMultipleEventsToICS
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Email();
            //ExStart: WriteMultipleEventsToICS
            IcsSaveOptions saveOptions = new IcsSaveOptions();
            saveOptions.Action = AppointmentAction.Create;
            using (CalendarWriter writer = new CalendarWriter(dataDir + "WriteMultipleEventsToICS_out.ics", saveOptions))
            {
                for (int i = 0; i < 10; i++)
                {
                    Appointment app = new Appointment(string.Empty, DateTime.Now, DateTime.Now, "sender@domain.com", "receiver@domain.com");
                    app.Description = "Test body " + i;
                    app.Summary = "Test summary:" + i;
                    writer.Write(app);
                }
            }
            //ExEnd: WriteMultipleEventsToICS
        }
    }
}
