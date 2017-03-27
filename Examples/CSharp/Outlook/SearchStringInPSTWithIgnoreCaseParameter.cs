using System;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System.IO;
using Aspose.Email.Tools.Search;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SearchStringInPSTWithIgnoreCaseParameter
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:SearchStringInPSTWithIgnoreCaseParameter
            string dataDir = RunExamples.GetDataDir_Outlook();

            string path = dataDir + "SearchStringInPSTWithIgnoreCaseParameter_out.pst";

            if (File.Exists(path))
            {
                File.Delete(path);
            }


            using (PersonalStorage personalStorage = PersonalStorage.Create(dataDir + "SearchStringInPSTWithIgnoreCaseParameter_out.pst", FileFormatVersion.Unicode))
            {
                FolderInfo folderInfo = personalStorage.CreatePredefinedFolder("Inbox", StandardIpmFolder.Inbox);

                folderInfo.AddMessage(MapiMessage.FromMailMessage(MailMessage.Load(dataDir + "Message.eml")));

                PersonalStorageQueryBuilder builder = new PersonalStorageQueryBuilder();
                // IgnoreCase is True
                builder.From.Contains("automated", true);

                MailQuery query = builder.GetQuery();
                MessageInfoCollection coll = folderInfo.GetContents(query);
                Console.WriteLine(coll.Count);
            }
            // ExEnd:SearchStringInPSTWithIgnoreCaseParameter
        }
    }
}