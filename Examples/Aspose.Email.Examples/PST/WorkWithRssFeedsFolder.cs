// Demonstrates how to create and then look up the standard RSS Feeds folder of a
// PST. Predefined folders are matched by their role, not by their display name.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class WorkWithRssFeedsFolder
    {
        public static void Run()
        {
            var pstPath = Data.Out/"WorkWithRssFeedsFolder_out.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                var created = pst.CreatePredefinedFolder("RSS Feeds", StandardIpmFolder.RssFeeds);
                Console.WriteLine($"Created:   {created.DisplayName} ({created.ContainerClass})");
            }

            using (var pst = PersonalStorage.FromFile(pstPath, false))
            {
                // The lookup is by role, so it keeps working even if the folder was
                // renamed in a localised Outlook.
                var rssFolder = pst.GetPredefinedFolder(StandardIpmFolder.RssFeeds);
                Console.WriteLine($"Retrieved: {rssFolder.DisplayName} ({rssFolder.ContainerClass})");
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }
    }
}
