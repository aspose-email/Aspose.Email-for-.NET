using System.IO;

namespace Aspose.Email.Live.Demos.UI.Models
{
    /// <summary>
    /// File processing result.
    /// </summary>
    public class FileSafeResult : BaseResult
	{
		/// <summary>
		/// Upload id.
		/// </summary>
		public string id { get; set; }
		
		/// <summary>
		/// File name.
		/// </summary>
		public string FileName { get; set; }

		/// <summary>
		/// Gets or sets the name of the upload file.
		/// </summary>
		public string UploadFileName { get; set; }

		/// <summary>
		/// idUpload from request.		
		/// </summary>
		public string idUpload { get; set; }

		/// <summary>
		/// FileSafeResult constructor.
		/// Sets IsSuccess to true.
		/// </summary>
		public FileSafeResult()
		{
			this.IsSuccess = true;
		}
	}
}
