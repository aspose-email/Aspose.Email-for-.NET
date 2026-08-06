// Demonstrates how to write a calendar file out in both ICS and MSG form, using the
// save-options type that matches each format.

using System;
using System.IO;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.Email
{
    internal static class SaveAppointmentToMsg
    {
        public static void Run()
        {
            var appointment = Appointment.Load(Data.Email/"test.ics");
            Console.WriteLine($"Appointment: {appointment.Summary}");

            var icsPath = Data.Out/"SaveAppointmentToMsg_out.ics";
            appointment.Save(icsPath, new AppointmentIcsSaveOptions());
            Console.WriteLine($"ICS: {new FileInfo(icsPath).Length} byte(s) -> {icsPath}");

            var msgPath = Data.Out/"SaveAppointmentToMsg_out.msg";
            appointment.Save(msgPath, new AppointmentMsgSaveOptions());
            Console.WriteLine($"MSG: {new FileInfo(msgPath).Length} byte(s) -> {msgPath}");
        }
    }
}
