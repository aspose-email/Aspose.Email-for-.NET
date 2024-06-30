using Aspose.Email;
using Aspose.Email.Storage.Pst;

namespace MailboxExtractor;

public class Pst(string pstFileName) : Mailbox(pstFileName)
{
    /// <summary>
    /// Override the SaveTo method to save messages from an MBOX file
    /// </summary>
    /// <param name="outputDir"></param>
    public override void SaveTo(string outputDir)
    {
        // Load the PST file
        using var pst = PersonalStorage.FromFile(FileName);
        // Process all folders
        ProcessPstFolder(pst.RootFolder, outputDir);
        
        Console.WriteLine();
        Console.WriteLine("PST extraction completed successfully.");
    }

    /// <summary>
    /// Recursive method to process PST folders and save messages
    /// </summary>
    /// <param name="folder"></param>
    /// <param name="folderDir"></param>
    private void ProcessPstFolder(FolderInfo folder, string folderDir)
    {
        // Create a directory for the current folder
        var currentFolderDir = Path.Combine(folderDir, SanitizeFolderName(folder.DisplayName));
    
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
        foreach (var subFolder in folder.GetSubFolders())
        {
            ProcessPstFolder(subFolder, currentFolderDir);
        }
    }
}