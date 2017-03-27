using System;
using System.IO;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class MergeFolderFromAnotherPSTFile
    {
        public static int totalAdded = 0;

        public static void Run()
        {
            // The path to the File directory.
            // ExStart:MergePSTFolders
            string dataDir = RunExamples.GetDataDir_Outlook();
            try
            {
                using (PersonalStorage destinationPst = PersonalStorage.FromFile(dataDir + @"destination.pst"))
                using (PersonalStorage sourcePst = PersonalStorage.FromFile(dataDir + @"source.pst"))
                {
                    FolderInfo destinationFolder = destinationPst.RootFolder.AddSubFolder("FolderFromAnotherPst");
                    FolderInfo sourceFolder = sourcePst.GetPredefinedFolder(StandardIpmFolder.DeletedItems);

                    // The events subscription is an optional step for the tracking process only.
                    destinationFolder.ItemMoved += destinationFolder_ItemMoved;

                    // Merges with the folder from another pst.
                    destinationFolder.MergeWith(sourceFolder);
                    Console.WriteLine("Total messages added: {0}", totalAdded);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\nThis example will only work if you apply a valid Aspose Email License. You can purchase full license or get 30 day temporary license from http:// Www.aspose.com/purchase/default.aspx.");
            }
            // ExEnd:MergePSTFolders
        }

        // ExStart:HelperMethods
        static void destinationFolder_ItemMoved(object sender, ItemMovedEventArgs e)
        {
            totalAdded++;
        }
        // ExEnd:HelperMethods

    }
}