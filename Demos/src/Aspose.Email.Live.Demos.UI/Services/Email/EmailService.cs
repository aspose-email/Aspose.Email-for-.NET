using Aspose.Email;
using System;
using System.IO;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	/// <summary>
	/// Class contains business logic for Email web apps
	/// </summary>
	public partial class EmailService : BaseService, IDisposable
    {
		/// <summary>
		/// Releases all resources.
		/// </summary>
		public void Dispose()
		{
			// Supported for future
		}

		static EmailService()
		{
			Models.License.SetAsposeEmailLicense();
			Models.License.SetAsposeSlidesLicense();
			Models.License.SetAsposeWordsLicense();
		}
	}
}
