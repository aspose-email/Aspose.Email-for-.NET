// Demonstrates how to create a new PST file, add an Inbox subfolder, and add a message to it.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddMessagesToPstFiles
    {
        public static void Run()
        {
            var path = Data.Out/"AddMessagesToPSTFiles_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                // AddSubFolder returns the new folder, so a second lookup is not needed.
                var inboxFolder = pst.RootFolder.AddSubFolder("Inbox");

                var message = MapiMessage.Load(Data.Mapi/"MapiMsgWithPoll.msg");
                inboxFolder.AddMessage(message);

                Console.WriteLine($"Added \"{message.Subject}\" to the {inboxFolder.DisplayName} folder.");
                Console.WriteLine($"Saved to {path}");
            }
        }
    }
}
