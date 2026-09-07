// Demonstrates the time zone an Outlook appointment carries.
//
// A MapiCalendarTimeZone is more than a name: it holds one rule per year the meeting
// could recur into, because the daylight-saving boundaries move. Each rule gives the
// offsets in minutes and the two switch-over dates, expressed as "the last Sunday in
// October" rather than a fixed date.

using System;
using System.Linq;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadCalendarTimeZoneRules
    {
        public static void Run()
        {
            var calendar = (MapiCalendar)MapiMessage.Load(Data.Mapi/"Test Meeting.msg").ToMapiMessageItem();

            Console.WriteLine($"Meeting: {calendar.Subject}");
            Console.WriteLine($"Starts:  {calendar.StartDate:u}");
            Console.WriteLine($"Ends:    {calendar.EndDate:u}");

            Print("Start", calendar.StartDateTimeZone);
            Print("End", calendar.EndDateTimeZone);
        }

        private static void Print(string label, MapiCalendarTimeZone timeZone)
        {
            Console.WriteLine($"\n{label} time zone: {(timeZone == null ? "(none)" : timeZone.KeyName)}");

            if (timeZone == null || timeZone.TimeZoneRules == null)
                return;

            foreach (var rule in timeZone.TimeZoneRules.Cast<MapiCalendarTimeZoneInfo>())
            {
                Console.WriteLine($"  rule for {rule.Year}");

                // Bias is the base offset from UTC in minutes, inverted: -300 is UTC+5.
                Console.WriteLine($"    base offset:     {-rule.Bias} minute(s) from UTC");
                Console.WriteLine($"    standard bias:   {rule.StandardBias}");
                Console.WriteLine($"    daylight bias:   {rule.DaylightBias}");
                Console.WriteLine($"    flags:           {rule.TimeZoneFlags}");

                PrintRule("    switches to standard", rule.StandardDate);
                PrintRule("    switches to daylight", rule.DaylightDate);
            }
        }

        private static void PrintRule(string label, MapiCalendarTimeZoneRule rule)
        {
            // Month 0 means the zone has no daylight-saving transition at all.
            if (rule == null || rule.Month == 0)
            {
                Console.WriteLine($"{label}: never");
                return;
            }

            Console.WriteLine($"{label}: month {rule.Month}, the {rule.Position} {rule.DayOfWeek} " +
                              $"at {rule.Hour:00}:{rule.Minute:00}");
        }
    }
}
