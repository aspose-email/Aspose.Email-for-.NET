using System.IO;
using System;
using Aspose.Email;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class NewPSTAddSubfolders
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "PersonalStorage.pst";

            if (File.Exists(dst))
                File.Delete(dst);
            
            // Create new PST
            PersonalStorage pst = PersonalStorage.Create(dst, FileFormatVersion.Unicode);

            // Add new folder "Inbox"
            pst.RootFolder.AddSubFolder("Inbox");

            Console.WriteLine(Environment.NewLine + "PST saved successfully at " + dst);
        }
    }
}
