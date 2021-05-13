namespace Aspose.Email.Live.Demos.UI.Models
{
    public class EmailAnotherApp
	{
		public string ImageSource { get; set; }
		public string ImageAlt { get; set; }
		public string AppName { get; set; }
		public string Href { get; set; }

		public EmailAnotherApp(string appName, string locale)
			: base()
		{
			AppName = appName;

			Href = $"{appName}";
			ImageAlt = $"{AppName}";

			ImageSource = $"https://www.aspose.cloud/templates/asposeapp/images/products/logo/aspose_{appName.ToLowerInvariant()}-app.png";
		}
	}
}
