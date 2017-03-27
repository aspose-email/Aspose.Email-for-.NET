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
    class RetreiveParentFolderInformationFromMessageInfo
    {
        public static void Run()
        {
            // ExStart:RetreiveParentFolderInformationFromMessageInfo
            // Load Pst File
            string dataDir = RunExamples.GetDataDir_Outlook() + "PersonalStorage.pst";
            using (PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir))
            {
                foreach (FolderInfo folder in personalStorage.RootFolder.GetSubFolders())
                {
                    foreach (MessageInfo msg in folder.EnumerateMessages())
                    {
                        FolderInfo fi = personalStorage.GetParentFolder(msg.EntryId);
                    }
                }
            }
            // ExEnd:RetreiveParentFolderInformationFromMessageInfo
        }
    }
}
