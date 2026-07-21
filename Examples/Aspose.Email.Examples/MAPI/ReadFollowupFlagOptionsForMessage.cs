// Demonstrates how to read follow-up flag options from a MapiMessage.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadFollowupFlagOptionsForMessage
    {
        public static void Run()
        {
            var mapi = MapiMessage.Load(Data.Mapi/"message.msg");
            var options = FollowUpManager.GetOptions(mapi);

            Console.WriteLine($"Subject:      {mapi.Subject}");
            Console.WriteLine($"Flag request: {Describe(options.FlagRequest)}");
            Console.WriteLine($"Start date:   {Describe(options.StartDate)}");
            Console.WriteLine($"Due date:     {Describe(options.DueDate)}");
            Console.WriteLine($"Reminder:     {Describe(options.ReminderTime)}");
            Console.WriteLine($"Completed:    {options.IsCompleted}");
        }

        // Unset follow-up dates come back as DateTime.MinValue rather than null,
        // so report them as "(not set)" instead of printing a year-1 timestamp.
        private static string Describe(DateTime value) =>
            value == DateTime.MinValue ? "(not set)" : value.ToString("u");

        private static string Describe(string value) =>
            string.IsNullOrEmpty(value) ? "(not set)" : value;
    }
}
