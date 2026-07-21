// Demonstrates how to create a new PST file and add a subfolder to it.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class NewPstAddSubfolders
    {
        public static void Run()
        {
            string dst = Data.Out/"PersonalStorage.pst";
            if (File.Exists(dst))
                File.Delete(dst);

            using (PersonalStorage pst = PersonalStorage.Create(dst, FileFormatVersion.Unicode))
            {
                pst.RootFolder.AddSubFolder("Inbox");
                Console.WriteLine("PST saved successfully at " + dst);
            }
        }
    }
}
