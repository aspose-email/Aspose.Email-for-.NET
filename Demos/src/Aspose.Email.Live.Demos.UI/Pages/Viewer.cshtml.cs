using Aspose.Email.Live.Demos.UI.Config;
using System;
using System.Web;

namespace Aspose.Email.Live.Demos.UI.Pages
{
    public class ViewerModel : BaseRootPageModel
	{
		public override string AppName => "Viewer";

		public string FileName { get; private set; }
		public string DownloadFileName { get; private set; }
		public string FileFormat { get; private set; }
		public string FolderName { get; private set; }
		public string CallbackURL { get; private set; }
		public string ApiUrl { get; private set; }

		public override void OnGet()
        {
			base.OnGet();

			if (!this.IsPostBack())
			{
				FileName = Get("filename");

				if (FileName.HasValue())
					DownloadFileName = HttpUtility.UrlEncode(FileName);

				FileFormat = Get("fileformat");
				FolderName = Get("foldername");
				CallbackURL = Get("callbackURL");

				ApiUrl = Configuration.WebApiPath;
			}
		}

		private string Get(string key)
		{
			return Request.RouteValues[key]?.ToString() ?? Request.Query[key];
		}

		public string GetAppRoute(string appName)
		{
			return "email/" + appName;
		}
	}
}
