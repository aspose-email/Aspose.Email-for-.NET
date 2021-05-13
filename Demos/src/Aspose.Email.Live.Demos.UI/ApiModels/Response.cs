using System.Collections.ObjectModel;

namespace Aspose.Email.Live.Demos.UI.Models
{
    ///<Summary>
    /// Response class to get or set any api call status
    ///</Summary>
    public class Response
	{
		///<Summary>
		/// Get or set DownloadFileLink
		///</Summary>
		public string DownloadFileLink { get; set; }
		///<Summary>
		/// Get or set StatusCode
		///</Summary>
		public int StatusCode { get; set; }
		///<Summary>
		/// Get or set FileName
		///</Summary>
		public string FileName { get; set; }
		///<Summary>
		/// Get or set FolderName
		///</Summary>
		public string FolderName { get; set; }

		///<Summary>
		/// DebugTrace
		///</Summary>
		public string Trace { get; set; }

		///<Summary>
		/// Get or set Status
		///</Summary>
		public string Status { get; set; }
		///<Summary>
		/// Get or set Text
		///</Summary>
		public string Text { get; set; }
		///<Summary>
		/// Get or set Files
		///</Summary>
		public Collection<string> Files;
		///<Summary>
		/// Get or set FileProcessingErrorCode
		///</Summary>
		public FileProcessingErrorCode FileProcessingErrorCode { get; set; }

		/// <summary>
		/// Get count of files
		/// </summary>
		public int FileCount { get { return Files != null ? Files.Count : 0; } }
	}
}
