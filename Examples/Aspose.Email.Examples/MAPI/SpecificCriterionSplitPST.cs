// Demonstrates how to split a PST file into several smaller PSTs, each holding the
// messages matching one search criterion - here, two consecutive date ranges.

using System;
using System.Collections.Generic;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SpecificCriterionSplitPST
    {
        public static void Run()
        {
            // One query per output PST: messages sent in the first week of April 2005,
            // then messages sent in the second week.
            var criteria = new List<MailQuery>();

            var firstWeek = new PersonalStorageQueryBuilder();
            firstWeek.SentDate.Since(new DateTime(2005, 04, 01));
            firstWeek.SentDate.Before(new DateTime(2005, 04, 07));
            criteria.Add(firstWeek.GetQuery());

            var secondWeek = new PersonalStorageQueryBuilder();
            secondWeek.SentDate.Since(new DateTime(2005, 04, 07));
            secondWeek.SentDate.Before(new DateTime(2005, 04, 13));
            criteria.Add(secondWeek.GetQuery());

            var outputDir = Data.OutSub("SplitPst");

            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage_New.pst"))
            {
                personalStorage.SplitInto(criteria, outputDir);
            }

            Console.WriteLine($"Split into {criteria.Count} PST file(s) in {outputDir}");
        }
    }
}
