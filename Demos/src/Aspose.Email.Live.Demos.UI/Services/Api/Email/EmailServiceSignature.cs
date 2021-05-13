using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email;
using Aspose.Email.Mapi;
using System;
using System.IO;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	public partial class EmailService
	{
		public void AddDrawingSignature(Stream input, string inputFileNameWithExtension, string imageBase64, string outputType, IOutputHandler handler)
		{
			PrepareOutputType(ref outputType);

			using (var mail = MapiHelper.GetMapiMessageFromStream(input))
			{
				AddDrawing(mail, imageBase64);

				if (outputType.Equals("MSG", StringComparison.OrdinalIgnoreCase))
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".signed.msg")))
						mail.Save(output);
				else
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".signed.eml")))
						mail.Save(output, new EmlSaveOptions(MailMessageSaveType.EmlFormat));
			}
		}

		public void AddTextSignature(Stream input, string inputFileNameWithExtension, string text, string textColor, string outputType, IOutputHandler handler)
		{
			PrepareOutputType(ref outputType);

			using (var mail = MapiHelper.GetMapiMessageFromStream(input))
			{
				AddText(mail, text, textColor);

				if (outputType.Equals("MSG", StringComparison.OrdinalIgnoreCase))
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".signed.msg")))
						mail.Save(output);
				else
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".signed.eml")))
						mail.Save(output, new EmlSaveOptions(MailMessageSaveType.EmlFormat));
			}
		}

		private MapiMessage AddDrawing(MapiMessage mail, string imageBase64)
		{
			var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");

			var element = htmlDocument.CreateElement("Signature");
			element.InnerHTML = "<img Src=\"data: image / png; base64, " + imageBase64 + "\" />";
			htmlDocument.Body.AppendChild(element);

			mail.SetBodyContent(htmlDocument.ToHTMLContentString(), BodyContentType.Html);
			return mail;
		}

		private MapiMessage AddText(MapiMessage mail, string text, string textColor)
		{
			var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");

			var element = htmlDocument.CreateElement("Signature");
			element.InnerHTML = "<p style = \"color:" + textColor + "\";>" + text + "</p>";

			htmlDocument.Body.AppendChild(element);

			mail.SetBodyContent(htmlDocument.ToHTMLContentString(), BodyContentType.Html);
			return mail;
		}
	}
}
