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
    class MoveItemsToOtherFolders
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:MoveItemsToOtherFolders
            string dataDir = RunExamples.GetDataDir_Outlook();
            using (PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir + "Outlook_1.pst"))
            {
                FolderInfo inbox = personalStorage.GetPredefinedFolder(StandardIpmFolder.Inbox);
                FolderInfo deleted = personalStorage.GetPredefinedFolder(StandardIpmFolder.DeletedItems);
                FolderInfo subfolder = inbox.GetSubFolder("Inbox");

                if (subfolder != null)
                {
                    // Move folder to the Deleted Items
                    personalStorage.MoveItem(subfolder, deleted);

                    // Move message to the Deleted Items
                    MessageInfoCollection contents = subfolder.GetContents();
                    personalStorage.MoveItem(contents[0], deleted);

                    // Move all inbox subfolders to the Deleted Items
                    inbox.MoveSubfolders(deleted);

                    // Move all subfolder contents to the Deleted Items
                    subfolder.MoveContents(deleted);
                }
            }
            // ExEnd:MoveItemsToOtherFolders
        }
    }
}