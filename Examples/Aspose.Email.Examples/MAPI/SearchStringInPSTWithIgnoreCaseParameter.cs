// Demonstrates how to search a PST folder with a case-insensitive query, using the
// ignoreCase parameter of the PersonalStorageQueryBuilder string comparisons.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SearchStringInPSTWithIgnoreCaseParameter
    {
        public static void Run()
        {
            var pstPath = Data.Out/"SearchStringInPSTWithIgnoreCaseParameter_out.pst";

            using (var personalStorage = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                var folderInfo = personalStorage.CreatePredefinedFolder("Inbox", StandardIpmFolder.Inbox);
                folderInfo.AddMessage(MapiMessage.FromMailMessage(MailMessage.Load(Data.Mapi/"Message.eml")));

                var builder = new PersonalStorageQueryBuilder();

                // The message was sent by "Automated Email Message", spelled with a capital
                // A. Passing ignoreCase: true makes the lowercase term match it anyway.
                builder.From.Contains("automated", true);

                var matches = folderInfo.GetContents(builder.GetQuery());
                Console.WriteLine($"Case-insensitive search for \"automated\" matched {matches.Count} message(s).");
            }
        }
    }
}
