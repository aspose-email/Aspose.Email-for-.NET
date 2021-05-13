using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using Microsoft.Extensions.Logging;
using System;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	/// <summary>
	/// Class contains business logic for Email web apps
	/// </summary>
	public partial class EmailService : IEmailService
	{
        public ILogger<EmailService> Logger { get; }

		public EmailService(ILogger<EmailService> logger)
        {
			Models.License.SetAsposeEmailLicense();
			Models.License.SetAsposeSlidesLicense();
			Models.License.SetAsposeWordsLicense();
			Models.License.SetAsposeCellsLicense();
			Models.License.SetAsposeHtmlLicense();
			Models.License.SetAsposeImagingLicense();

            Logger = logger;
        }

		void HandleFolderAndSubfolders(Action<MapiMessage> handler, FolderInfo folderInfo)
		{
			foreach (MapiMessage mapiMessage in folderInfo.EnumerateMapiMessages())
			{
				handler(mapiMessage);
			}

			if (folderInfo.HasSubFolders == true)
			{
				foreach (FolderInfo subfolderInfo in folderInfo.GetSubFolders())
				{
					HandleFolderAndSubfolders(handler, subfolderInfo);
				}
			}
		}
	}
}
