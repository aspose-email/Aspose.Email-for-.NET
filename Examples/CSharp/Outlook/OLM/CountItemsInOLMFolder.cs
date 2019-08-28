using System;
using Aspose.Email.Storage.Olm;
using Aspose.Email.Mapi;
using System.Collections.Generic;

namespace Aspose.Email.Examples.CSharp.Email.Outlook.OLM
{
    class CountItemsInOLMFolder
    {
        // ExStart:1
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            OlmStorage storage = new OlmStorage(dataDir + "SampleOLM.olm");
            PrintMessageCount(storage.FolderHierarchy);
        }

        public static void PrintMessageCount(List<OlmFolder> folders)
        {
            foreach (OlmFolder folder in folders)
            {
                Console.WriteLine("Message Count [" + folder.Name + "]: " + folder.MessageCount);
            }
        }
        // ExEnd:1
    }
}
