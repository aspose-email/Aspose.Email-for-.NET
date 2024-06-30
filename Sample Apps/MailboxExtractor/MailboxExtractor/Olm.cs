using Aspose.Email;
using Aspose.Email.Storage.Olm;

namespace MailboxExtractor;

public class Olm(string olmFileName) : Mailbox(olmFileName)
{
    /// <summary>
    /// Override the SaveTo method to save messages from an OLM file.
    /// </summary>
    /// <param name="outputDir"></param>
    public override void SaveTo(string outputDir)
    {
        // Open the OLM file
        using var olm = OlmStorage.FromFile(FileName);
        // Process the folders and save messages  
        ProcessOlmFolder(olm.GetFolders(), outputDir);
        
        Console.WriteLine();
        Console.WriteLine("OLM extraction completed successfully.");
    }
    
    /// <summary>
    /// Recursive method to process OLM folders and save messages
    /// </summary>
    /// <param name="folderHierarchy"></param>
    /// <param name="folderDir"></param>
    private void ProcessOlmFolder(List<OlmFolder> folderHierarchy, string folderDir)
    {
        foreach (var folder in folderHierarchy)
        {
            // Create a directory for the current folder
            var currentFolderDir = Path.Combine(folderDir, SanitizeFolderName(folder.Name));
        
            // Create the directory if it doesn't exist
            if (!Directory.Exists(currentFolderDir))
            {
                Directory.CreateDirectory(currentFolderDir);  
            }
        
            // Process all messages in the current folder
            foreach (var message in folder.EnumerateMapiMessages())
            {
                var msgFilePath = Path.Combine(currentFolderDir, SanitizeFileName($"{message.Subject}.msg"));

                // Save the message in MSG format
                message.Save(msgFilePath, SaveOptions.DefaultMsgUnicode);
            }

            // Process subfolders recursively
            ProcessOlmFolder(folder.SubFolders, currentFolderDir);
        }
    }
}
