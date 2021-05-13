namespace Aspose.Email.Live.Demos.UI.Models
{
	public class FileDataResult
	{
		public string Html { get; set; }
		public FileAttachment[] Attachments { get; set; }
	}

	public class FileAttachment
	{
		public string Name { get; set; }
		public string Type { get; set; }
		public string Extension { get; set; }
		public string Size { get; set; }
		public byte[] Source { get; set; }
	}
}
