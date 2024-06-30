using Aspose.Email;
using Aspose.Email.Tools;

namespace MailboxExtractor;

/// <summary>
/// Base class.
/// </summary>
/// <param name="fileName"></param>
public abstract class Mailbox(string fileName)
{
    protected readonly string FileName = fileName;

    /// <summary>
    /// Abstract method to save messages to the output directory
    /// </summary>
    /// <param name="outputDir">The output directory.</param>
    public abstract void SaveTo(string outputDir);
    
    /// <summary>
    /// Factory method to open a mailbox file and return the appropriate Mailbox type
    /// </summary>
    /// <param name="fileName">Name of a storage file.</param>
    /// <returns></returns>
    /// <exception cref="FormatNotSupportedException">Unsupported file format.</exception>
    public static Mailbox Open(string fileName)
    {
        var info = FileFormatUtil.DetectFileFormat(fileName);
    
        Console.WriteLine($"File format: {info.FileFormatType}");
        Console.WriteLine();

        switch (info.FileFormatType)
        {
            case FileFormatType.Ost:
            case FileFormatType.Pst:
                return new Pst(fileName);
            case FileFormatType.Mbox:
                return new Mbox(fileName);
            case FileFormatType.Olm:
                return new Olm(fileName);
            case FileFormatType.Tgz:
                return new Tgz(fileName);
            default:
                throw new FormatNotSupportedException("Unsupported file format.");
        }
    }

    /// <summary>
    /// Sanitize folder name by replacing invalid path characters with '_'.
    /// </summary>
    /// <param name="folderName">A folder name.</param>
    /// <returns></returns>
    protected static string SanitizeFolderName(string folderName)
    {
        foreach (var c in Path.GetInvalidPathChars())
        {
            folderName = folderName.Replace(c, '_');
        }
    
        return folderName;
    }

    /// <summary>
    /// Sanitize file name by replacing invalid file name characters with '_'.
    /// </summary>
    /// <param name="fileName">Name of a message file.</param>
    /// <returns></returns>
    protected static string SanitizeFileName(string fileName)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(c, '_');
        }
    
        return fileName;
    }
}
