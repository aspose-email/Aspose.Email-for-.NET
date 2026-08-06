// Demonstrates how to find out what kind of Outlook item a MapiMessage really is,
// so it can be cast to the right type before its specific properties are read.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetMapiItemType
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage.pst", false))
            {
                Walk(pst, pst.RootFolder);
            }
        }

        private static void Walk(PersonalStorage pst, FolderInfo folder)
        {
            foreach (var messageInfo in folder.EnumerateMessages())
            {
                var msg = pst.ExtractMessage(messageInfo);
                Console.Write($"{msg.Subject}: {msg.SupportedType} -> ");

                switch (msg.SupportedType)
                {
                    case MapiItemType.Contact:
                        var contact = (MapiContact)msg.ToMapiMessageItem();
                        Console.WriteLine(contact.NameInfo.DisplayName);
                        break;
                    case MapiItemType.Calendar:
                        var calendar = (MapiCalendar)msg.ToMapiMessageItem();
                        Console.WriteLine($"{calendar.StartDate} - {calendar.EndDate}");
                        break;
                    case MapiItemType.DistList:
                        var dl = (MapiDistributionList)msg.ToMapiMessageItem();
                        Console.WriteLine($"{dl.Members.Count} member(s)");
                        break;
                    case MapiItemType.Journal:
                        var journal = (MapiJournal)msg.ToMapiMessageItem();
                        Console.WriteLine(journal.BriefDescription);
                        break;
                    case MapiItemType.Note:
                        var note = (MapiNote)msg.ToMapiMessageItem();
                        Console.WriteLine(note.Color);
                        break;
                    case MapiItemType.Task:
                        var task = (MapiTask)msg.ToMapiMessageItem();
                        Console.WriteLine(task.Status);
                        break;
                    default:
                        Console.WriteLine(msg.SenderName);
                        break;
                }
            }

            foreach (var subFolder in folder.GetSubFolders())
                Walk(pst, subFolder);
        }
    }
}
