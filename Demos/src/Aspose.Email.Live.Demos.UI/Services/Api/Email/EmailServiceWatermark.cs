using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Html;
using Aspose.Imaging;
using Aspose.Imaging.Brushes;
using Aspose.Imaging.ImageOptions;
using Aspose.Imaging.Sources;
using System;
using System.IO;
using Graphics = Aspose.Imaging.Graphics;
using Image = Aspose.Imaging.Image;
using System.Linq;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	public partial class EmailService
	{
		public void AddTextWatermark(Stream input, string inputFileNameWithExtension, string watermarkText, string fontFamily, string watermarkColor, IOutputHandler handler)
		{
			using (var mail = MapiHelper.GetMapiMessageFromStream(input, out FileFormatInfo formatInfo))
			{
				var html = mail.BodyHtml;

				var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");
				var options = new PngOptions();
				options.Source = new StreamSource(new MemoryStream(), true);

				var width = watermarkText.Length * 17;
				var height = 120;

				using (Image image = Image.Create(options, width, height))
				{
					// Create graphics object to perform draw operations.
					var graphics = new Graphics(image);

					// Create font to draw watermark with.
					var font = new Aspose.Imaging.Font(fontFamily, 20.0f);

					// Create a solid brush with color alpha set near to 0 to use watermarking effect.
					using (SolidBrush backgroundBrush = new SolidBrush(Color.White))
					{
						using (SolidBrush brush = new SolidBrush(GetColor(watermarkColor)))
						{
							// Specify string alignment to put watermark at the image center.
							StringFormat sf = new StringFormat();
							sf.Alignment = StringAlignment.Center;
							sf.LineAlignment = StringAlignment.Center;

							graphics.FillRectangle(backgroundBrush, new Rectangle(0, 0, image.Width, image.Height));
							// Draw watermark using font, partly-transparent brush and rotation matrix at the image center.
							graphics.DrawString(watermarkText, font, brush, new RectangleF(0, 0, image.Width, image.Height), sf);
						}
					}

					var ms = new MemoryStream();
					image.Save(ms);
					mail.Attachments.Add("watermark", ms.ToArray());
				}

				var attachment = mail.Attachments.Find(x => x.LongFileName == "watermark");
				attachment.SetContentId("watermark");

				var bodyHtml = htmlDocument.Body.InnerHTML;

				var watermarkHtml = $@"<table role=""presentation"" style=""width:100%; height:1200px; background-image:url(cid:watermark); background-repeat:repeat; table-layout: fixed;"" cellpadding=""0"" cellspacing=""0"" border=""0"">
							<tr>
								<td valign=""top"" style=""width:100%; height:1200px; background-size:cover; background-image:url(cid:watermark); background-repeat:repeat; word-break: normal; border-collapse : collapse !important;"">

									<!--[if gte mso 9]>
									<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" filltype=""tile"" stroke=""false"" style=""height:1200px; width:640px; border: 0;display: inline-block;position: absolute; background-image:url(cid:watermark); background-repeat:repeat; word-break: normal; border-collapse : collapse !important; mso-position-vertical-relative:text;"">
									<v:fill type=""tile"" opacity=""100%"" src=""cid:watermark"" />
									<v:textbox inset=""0,0,0,0"">
									<![endif]-->

									<div>
										{bodyHtml}
									</div>

									<!--[if gte mso 9]>\
									</v:textbox>
									</v:fill>
									</v:rect>
									<![endif]-->

								</td>
							</tr>
						</table>";

				htmlDocument.Body.InnerHTML = watermarkHtml;

				var content = htmlDocument.ToHTMLContentString();

				mail.SetBodyContent(content, BodyContentType.Html);

				if (formatInfo.FileFormatType == FileFormatType.Msg)
				{
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".marked.msg")))
						mail.Save(output, SaveOptions.DefaultMsgUnicode);
				}
				else
				{
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".marked.eml")))
						mail.Save(output, SaveOptions.DefaultEml);
				}
			}
		}

		public void AddImageWatermark(Stream input, string inputFileNameWithExtension, byte[] imageBytes, IOutputHandler handler)
		{
			using (var mail = MapiHelper.GetMapiMessageFromStream(input, out FileFormatInfo formatInfo))
			{
				mail.Attachments.Add("watermark", imageBytes);

				var html = mail.BodyHtml;
				var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");

				var attachment = mail.Attachments.Find(x => x.LongFileName == "watermark");
				attachment.SetContentId("watermark");

				var bodyHtml = htmlDocument.Body.InnerHTML;

				var watermarkHtml = $@"<table role=""presentation"" style=""width:100%; height:1200px; background-image:url(cid:watermark); background-repeat:repeat; table-layout: fixed;"" cellpadding=""0"" cellspacing=""0"" border=""0"">
							<tr>
								<td valign=""top"" style=""width:100%; height:1200px; background-size:cover; background-image:url(cid:watermark); background-repeat:repeat; word-break: normal; border-collapse : collapse !important;"">

									<!--[if gte mso 9]>
									<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" filltype=""tile"" stroke=""false"" style=""height:1200px; width:640px; border: 0;display: inline-block;position: absolute; background-image:url(cid:watermark); background-repeat:repeat; word-break: normal; border-collapse : collapse !important; mso-position-vertical-relative:text;"">
									<v:fill type=""tile"" opacity=""100%"" src=""cid:watermark"" />
									<v:textbox inset=""0,0,0,0"">
									<![endif]-->

									<div>
										{bodyHtml}
									</div>

									<!--[if gte mso 9]>\
									</v:textbox>
									</v:fill>
									</v:rect>
									<![endif]-->

								</td>
							</tr>
						</table>";

				htmlDocument.Body.InnerHTML = watermarkHtml;

				var content = htmlDocument.ToHTMLContentString();

				mail.SetBodyContent(content, BodyContentType.Html);

				if (formatInfo.FileFormatType == FileFormatType.Msg)
				{
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".marked.msg")))
						mail.Save(output, SaveOptions.DefaultMsgUnicode);
				}
				else
				{
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".marked.eml")))
						mail.Save(output, SaveOptions.DefaultEml);
				}
			}
		}

		public void RemoveWatermark(Stream input, string inputFileNameWithExtension, IOutputHandler handler)
		{
			using (var mail = MapiHelper.GetMapiMessageFromStream(input, out FileFormatInfo formatInfo))
			{
				var html = mail.BodyHtml;
				var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");

				var attachment = mail.Attachments.Find(x => x.LongFileName.Equals("watermark", StringComparison.OrdinalIgnoreCase));

				if (attachment != null)
					mail.Attachments.Remove(attachment);

				// Remove body backgrounds
				htmlDocument.Body.RemoveAttribute("background");
				htmlDocument.Body.Style.Background = string.Empty;
				htmlDocument.Body.Style.BackgroundImage = string.Empty;

				HTMLElement element = htmlDocument.Body;

				// Remove background from all single elements.
				do
				{
					// Skip unnecessary elements
					element = (HTMLElement)element.Children.Where(x => !(x is HTMLTitleElement) && !(x is HTMLStyleElement)).ElementAt(0);

					// Process only default tags for watermark
					if (element.NodeName != "TABLE" &&
						element.NodeName != "TITLE" &&
						element.NodeName != "TR" &&
						element.NodeName != "TD" &&
						element.NodeName != "TBODY")
					{
						// If reached div, content starting here
						if (element.NodeName == "DIV")
							htmlDocument.Body.InnerHTML = element.InnerHTML;

						break;
					}

					// Remove comments
					for (int i = 0; i < element.ChildNodes.Length; i++)
					{
						var node = element.ChildNodes[i];

						if (node.NodeType == Html.Dom.Node.COMMENT_NODE &&
							node.NodeValue != null &&
							node.NodeValue.StartsWith("[if gte mso 9]", StringComparison.OrdinalIgnoreCase))
						{
							element.RemoveChild(node);
							i--;
						}
					}

					element.Style.Background = string.Empty;
					element.Style.BackgroundImage = string.Empty;
				}
				while (element.ChildElementCount == 1);

				var content = htmlDocument.ToHTMLContentString();

				mail.SetBodyContent(content, BodyContentType.Html);

				if (formatInfo.FileFormatType == FileFormatType.Msg)
				{
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".unmarked.msg")))
						mail.Save(output, SaveOptions.DefaultMsgUnicode);
				}
				else
				{
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".unmarked.eml")))
						mail.Save(output, SaveOptions.DefaultEml);
				}
			}
		}

		private Color GetColor(string watermarkColor)
		{
			Color color;
			if (!watermarkColor.IsNullOrWhiteSpace())
			{
				if (!watermarkColor.StartsWith("#"))
				{
					watermarkColor = "#" + watermarkColor;
				}
				System.Drawing.Color sColor = System.Drawing.ColorTranslator.FromHtml(watermarkColor);
				color = Color.FromArgb(sColor.A, sColor.R, sColor.G, sColor.B);
			}
			else
			{
				color = Color.FromArgb(50, 128, 128, 128);
			}
			return color;
		}
	}
}
