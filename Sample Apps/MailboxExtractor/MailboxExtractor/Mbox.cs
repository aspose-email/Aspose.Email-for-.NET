using Aspose.Email;
using Aspose.Email.Storage.Mbox;

namespace MailboxExtractor;

public class Mbox(string mboxFileName) : Mailbox(mboxFileName)
{
    /// <summary>
    /// Override the SaveTo method to save messages from an MBOX file
    /// </summary>
    /// <param name="outputDir"></param>
    public override void SaveTo(string outputDir)
    {
        // Open the MBOX file
        using var mbox = MboxStorageReader.CreateReader(FileName, new MboxLoadOptions());  
    
        // Create a directory for the MBOX file
        var currentFolderDir = Path.Combine(outputDir, "Mbox");
        
        // Create the directory if it doesn't exist
        if (!Directory.Exists(currentFolderDir))
        {
            Directory.CreateDirectory(currentFolderDir);  
        }

        // Process all messages in the MBOX file
        foreach (var message in mbox.EnumerateMessages())
        {
            var msgFilePath = Path.Combine(currentFolderDir, SanitizeFileName($"{message.Subject}.eml"));
        
            // Save the message in EML format
            message.Save(msgFilePath, SaveOptions.DefaultEml);
        }
        
        Console.WriteLine();
        Console.WriteLine("MBOX extraction completed successfully.");
    }
}
