// Demonstrates how to read items from an Outlook PST folder and save each one
// to disk in iCalendar (ICS) format.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveCalendarItems
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("Calendar");

            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst"))
            {
                var folderInfo = pst.RootFolder.GetSubFolder("Inbox");
                var saved = 0;

                foreach (var messageInfo in folderInfo.GetContents())
                {
                    var calendar = (MapiMessage)pst.ExtractMessage(messageInfo).ToMapiMessageItem();

                    Console.WriteLine($"Name: {calendar.Subject}");

                    calendar.Save(outputDir/(calendar.Subject + "_out.ics"));
                    saved++;
                }

                Console.WriteLine($"\nSaved {saved} item(s) as ICS to {outputDir}");
            }
        }
    }
}
