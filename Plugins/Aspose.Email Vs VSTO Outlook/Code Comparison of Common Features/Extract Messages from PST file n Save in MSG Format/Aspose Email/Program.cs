using Aspose.Email.Outlook;
using Aspose.Email.Outlook.Pst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            string pstFilePath = "sample.pst";
            // Create an instance of PersonalStorage and load the PST from file
            using (PersonalStorage personalStorage = PersonalStorage.FromFile(pstFilePath))
            {
                // Get the list of subfolders in PST file
                FolderInfoCollection folderInfoCollection = personalStorage.RootFolder.GetSubFolders();
                // Traverse through all folders in the PST file
                // TODO: This is not recursive
                foreach (FolderInfo folderInfo in folderInfoCollection)
                {
                    // Get all messages in this folder
                    MessageInfoCollection messageInfoCollection = folderInfo.GetContents();
                    // Loop through all the messages in this folder
                    foreach (MessageInfo messageInfo in messageInfoCollection)
                    {
                        // Extract the message in MapiMessage instance
                        MapiMessage message = personalStorage.ExtractMessage(messageInfo);
                        Console.WriteLine("Saving message {0} ....", message.Subject);
                        // Save the message to disk in MSG format
                        // TODO: File name may contain invalid characters [\ / : * ? " < > |]
                        message.Save(@"\extracted\" + message.Subject + ".msg");
                    }
                }
            }
        }
    }
}
