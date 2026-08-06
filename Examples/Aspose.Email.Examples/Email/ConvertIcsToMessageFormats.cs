// Demonstrates the shortcut from a calendar file to the message formats: an
// Appointment converts straight to MailMessage and MapiMessage, no manual assembly.

using System;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.Email
{
    internal static class ConvertIcsToMessageFormats
    {
        public static void Run()
        {
            var appointment = Appointment.Load(Data.Email/"test.ics");
            Console.WriteLine($"Appointment: {appointment.Summary} ({appointment.StartDate} - {appointment.EndDate})");

            var eml = appointment.ToMailMessage();
            var emlPath = Data.Out/"ConvertIcsToMessageFormats_out.eml";
            eml.Save(emlPath, SaveOptions.DefaultEml);
            Console.WriteLine($"As MailMessage: {eml.Subject} -> {emlPath}");

            var msg = appointment.ToMapiMessage();
            var msgPath = Data.Out/"ConvertIcsToMessageFormats_out.msg";
            msg.Save(msgPath);
            Console.WriteLine($"As MapiMessage: {msg.Subject} -> {msgPath}");
        }
    }
}
