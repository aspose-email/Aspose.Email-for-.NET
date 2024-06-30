using Aspose.Email.Storage.Zimbra;

namespace MailboxExtractor;

public class Tgz(string tgzFileName) : Mailbox(tgzFileName)
{
    /// <summary>
    /// Override the SaveTo method to save messages from a TGZ file
    /// </summary>
    /// <param name="outputDir"></param>
    public override void SaveTo(string outputDir)
    {
        // Open the TGZ file
        using var tgz = new TgzReader(FileName);
        
        // Export the contents to the output directory  
        tgz.ExportTo(outputDir);  
        
        Console.WriteLine();
        Console.WriteLine("TGZ extraction completed successfully.");
    }
}
