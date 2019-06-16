using System;
using Aspose.Email.Storage.Olm;
using Aspose.Email.Mapi;
using System.Collections.Generic;

namespace Aspose.Email.Examples.CSharp.Email.Outlook.OLM
{
    class GetFolderPathInOLM
    {
        // ExStart:1
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            OlmStorage storage = new OlmStorage(dataDir + "SampleOLM.olm");
            PrintPath(storage, storage.FolderHierarchy);
        }

        public static void PrintPath(OlmStorage storage, List<OlmFolder> folders)
        {
            foreach (OlmFolder folder in folders)
            {
                // print the current folder path
                Console.WriteLine(folder.Path);

                if (folder.SubFolders.Count > 0)
                {
                    PrintPath(storage, folder.SubFolders);
                }
            }
        }
        // ExEnd:1
    }
}
