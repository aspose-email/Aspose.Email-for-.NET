using System;
using Aspose.Email.Storage.Olm;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.CSharp.Email.Outlook.OLM
{
    class LoadAndReadOLMFile
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "OutlookforMac.olm";
            // ExStart:LoadAndReadOLMFile
            using (OlmStorage storage = new OlmStorage(dst))
            {
                foreach (OlmFolder folder in storage.FolderHierarchy)
                {
                    if (folder.HasMessages)
                    {
                        // extract messages from folder
                        foreach (MapiMessage msg in storage.EnumerateMessages(folder))
                        {
                            Console.WriteLine("Subject: " + msg.Subject);
                        }
                    }

                    // read sub-folders
                    if (folder.SubFolders.Count > 0)
                    {
                        foreach (OlmFolder sub_folder in folder.SubFolders)
                        {
                            Console.WriteLine("Subfolder: " + sub_folder.Name);
                        }
                    }
                }
            }
            // ExEnd:LoadAndReadOLMFile
        }
    }
}
