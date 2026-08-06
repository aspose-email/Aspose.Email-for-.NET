// Demonstrates how to page through the results of a PST folder query, so a folder
// holding thousands of messages never has to be materialised in one go.

using System;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetPSTContentsWithPaging
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst", false))
            {
                var folder = pst.RootFolder.GetSubFolder("Inbox");

                var queryBuilder = new PersonalStorageQueryBuilder();
                queryBuilder.Subject.Contains("Test", true);
                var query = queryBuilder.GetQuery();

                const int pageSize = 5;

                for (var pageIndex = 0; ; pageIndex++)
                {
                    var startIndex = pageIndex * pageSize;

                    // GetContents(query, startIndex, count) reads only the requested slice.
                    var messages = folder.GetContents(query, startIndex, pageSize);
                    if (messages.Count == 0)
                        break;

                    Console.WriteLine($"--- page {pageIndex + 1} ({messages.Count} message(s)) ---");
                    foreach (var messageInfo in messages)
                        Console.WriteLine($"  {messageInfo.Subject} / {messageInfo.SenderRepresentativeName}");
                }
            }
        }
    }
}
