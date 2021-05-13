using Aspose.Html.Saving;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
	public static class HtmlHelper
	{
		public static string ToHTMLContentString(this Aspose.Html.HTMLDocument document)
		{
			var storage = new HtmlOutputInMemoryStorage();
			document.Save(storage, HTMLSaveFormat.HTML);
			return storage.ReadResourceAsString(0);
		}
	}
}

