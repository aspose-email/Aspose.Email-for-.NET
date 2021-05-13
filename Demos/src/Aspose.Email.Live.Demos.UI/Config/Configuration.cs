using Microsoft.Extensions.Configuration;
using System;

namespace Aspose.Email.Live.Demos.UI.Config
{
    public static class Configuration
	{
		public static IConfiguration Settings { get; private set; }
		public static void SetConfiguration(IConfiguration settings)
		{
			Settings = settings;

			var api = settings.GetValue<string>("WebApi:Path");
			var port = settings.GetValue<string>("WebApi:Port");

			if (!port.IsNullOrWhiteSpace())
				api = api + ":" + port;

			WebApiPath = api + "/";
			WebApiFileDownloadPath = api + settings.GetValue<string>("WebApi:FileDownloadPath");
		}

		public static string WebApiPath { get; private set; }
		public static string WebApiFileDownloadPath { get; private set; }
	}
}
