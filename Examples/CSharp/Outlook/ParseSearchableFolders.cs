using System.IO;
using System;
using System.Collections.Generic;
using Aspose.Email.Storage.Pst;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/


namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ParseSearchableFolders
    {
        // ExStart:ParseSearchableFolders
        public static void Run()
        {
            WalkAllFolders();            
        }

        // Open an OST and get the name and
        // parent name for all folders in the file
        private static void WalkAllFolders()
        {            
            string dataDir = RunExamples.GetDataDir_Outlook();
            string path = dataDir + "PersonalStorage.pst";
            byte[] buffer = File.ReadAllBytes(path);
            using (Stream s = File.OpenRead(path))
            {
                PersonalStorage pst = PersonalStorage.FromStream(s);

                // One sample of the sub-folder EntryId which could not be retrieved in Finder sub-folder
                FolderInfo target = pst.GetFolderById("AAAAAB9of1CGOidPhTb686WQY68igAAA");
                IList<string> folderData = new List<string>();
                WalkFolders(pst.RootFolder, "N/A", folderData);
            }            
        }

        // Walk the folder structure recursively
        private static void WalkFolders(FolderInfo folder, string parentFolderName, IList<string> folderData)
        {
            string displayName = (string.IsNullOrEmpty(folder.DisplayName)) ? "ROOT" : folder.DisplayName;
            string folderNames = string.Format("DisplayName = {0}; Parent.DisplayName = {1}", displayName, parentFolderName);
            folderData.Add(folderNames);
            if (displayName == "Finder")
            {
                Console.WriteLine("Test this case");
            }

            if (!folder.HasSubFolders)
            {
                return;
            }
            FolderInfoCollection coll = folder.GetSubFolders(FolderKind.Search | FolderKind.Normal);
            foreach (FolderInfo subfolder in coll)
            {
                WalkFolders(subfolder, displayName, folderData);
            }
        }
        //ExEnd:ParseSearchableFolders
    }
}