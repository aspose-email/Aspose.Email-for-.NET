using Aspose.Email.Live.Demos.UI.Models;
using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.Models
{
	public class EditorUpdateContentRequest
	{
		public string FileName { get; set; }
		public string FolderName { get; set; }
		public string Htmldata { get; set; }
		public string OutputType { get; set; }
		public bool CreateNewFile { get; set; }
		public Dictionary<int, FileAttachmentChange> AttachmentsData { get; set; }
	}
}
