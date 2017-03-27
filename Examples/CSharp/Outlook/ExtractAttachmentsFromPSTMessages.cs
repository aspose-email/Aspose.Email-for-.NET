using Aspose.Email.Mapi;
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
    class ExtractAttachmentsFromPSTMessages
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:ExtractAttachmentsFromPSTMessages
            string dataDir = RunExamples.GetDataDir_Outlook();

            using (PersonalStorage personalstorage = PersonalStorage.FromFile(dataDir + "Outlook.pst"))
            {
                FolderInfo folder = personalstorage.RootFolder.GetSubFolder("Inbox");

                foreach (var messageInfo in folder.EnumerateMessagesEntryId())
                {
                    MapiAttachmentCollection attachments = personalstorage.ExtractAttachments(messageInfo);

                    if (attachments.Count != 0)
                    {
                        foreach (var attachment in attachments)
                        {
                            if (!string.IsNullOrEmpty(attachment.LongFileName))
                            {
                                if (attachment.LongFileName.Contains(".msg"))
                                {
                                    continue;
                                }
                                else
                                {
                                    attachment.Save(dataDir + @"\Attachments\" + attachment.LongFileName);
                                }
                            }
                        }
                    }
                }
            }
            // ExEnd:ExtractAttachmentsFromPSTMessages
        }
    }
}