// Demonstrates how to recursively extract all messages from a PST file and save them as MSG.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ExtractMessagesFromPstFile
    {
        public static void Run()
        {
            try
            {
                using (var pst = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage.pst"))
                {
                    Console.WriteLine("Display Format: " + pst.Format);
                    ExtractMsgFiles(pst.RootFolder, pst);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void ExtractMsgFiles(FolderInfo folder, PersonalStorage pst)
        {
            Console.WriteLine("Folder: " + folder.DisplayName);
            Console.WriteLine("==================================");

            foreach (var info in folder.GetContents())
            {
                Console.WriteLine("Saving message {0} ....", info.Subject);
                var message = pst.ExtractMessage(info);
                message.Save(Data.Out/(message.Subject.Replace(":", " ") + ".msg"));

                using (var ms = new MemoryStream())
                    message.Save(ms);
            }

            foreach (var subfolder in folder.GetSubFolders())
                ExtractMsgFiles(subfolder, pst);
        }
    }
}
