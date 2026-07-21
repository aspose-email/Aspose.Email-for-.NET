// Demonstrates how to recursively display folder names and message subjects from a PST file.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetMessageInformation
    {
        public static void Run()
        {
            try
            {
                using (var pst = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage.pst"))
                {
                    Console.WriteLine("Display Format: " + pst.Format);
                    DisplayFolderContents(pst.RootFolder, pst);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void DisplayFolderContents(FolderInfo folder, PersonalStorage pst)
        {
            Console.WriteLine("Folder: " + folder.DisplayName);
            Console.WriteLine("==================================");

            foreach (var msg in folder.GetContents())
            {
                Console.WriteLine("Subject: " + msg.Subject);
                Console.WriteLine("Sender: " + msg.SenderRepresentativeName);
                Console.WriteLine("Recipients: " + msg.DisplayTo);
                Console.WriteLine("------------------------------");
            }

            foreach (var subfolder in folder.GetSubFolders())
                DisplayFolderContents(subfolder, pst);
        }
    }
}
