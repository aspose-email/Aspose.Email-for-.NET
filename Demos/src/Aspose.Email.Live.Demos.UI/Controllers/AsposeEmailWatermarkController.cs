using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Html;
using Aspose.Imaging;
using Aspose.Imaging.Brushes;
using Aspose.Imaging.ImageOptions;
using Aspose.Imaging.Sources;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    /// <summary>
    ///Implements the AsposeEmailWatermarkController Class to add or remove watermarks
    /// </summary>
    public class AsposeEmailWatermarkController : EmailBase
	{
		/// <summary>
		///Init the AsposeEmailWatermarkController Class 
		/// </summary>
		public AsposeEmailWatermarkController()
		{
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
		}

		[MimeMultipart]
		[HttpPost]
		[AcceptVerbs("GET", "POST")]
		public async Task<Response> ProcessWatermark(string type, string textColor = null, string fontFamily = "Arial")
		{
			(string folderName, string[] fileNames) = (null, null);

			try
			{
				(folderName, fileNames) = await UploadFiles();

				var text = ReadAndRemoveAsText(ref fileNames, folderName, "text");

				byte[] image = null;

				if (type == "image")
					image = ReadAndRemoveAsBytes(ref fileNames, folderName, fileNames.Last());

				var processor = new CustomSingleOrZipFileProcessor()
				{
					
					CustomFilesProcessMethod = (string[] inputFilePaths, string outputFolderPath) =>
					{
						for (int i = 0; i < inputFilePaths.Length; i++)
						{
							var filePath = inputFilePaths[i];

							var mail = MapiHelper.GetMapiMessageFromFile(filePath);

							if (mail == null)
								throw new Exception("Invalid file format");

							switch (type)
							{
								case "text":
									AddTextWatermark(filePath, mail, outputFolderPath, text, fontFamily, textColor);
									break;
								case "image":
									AddImageWatermark(filePath, mail, outputFolderPath, image);
									break;
								case "remove":
									RemoveWatermark(filePath, mail, outputFolderPath);
									break;
								default:
									break;
							}
						}
					}
				};


				return processor.Process(folderName, fileNames);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response() { Status = "200", Text = "Error on processing file" };
			}
		}

		/// <summary>
		///AddWatermarkToEmail method to add watermarks
		/// </summary>
		[HttpGet]
		[AcceptVerbs("GET", "POST")]
		public Task<Response> AddWatermarkToEmail(string fileName, string folderName, string watermarkText, string fontFamily, string outputType, string watermarkColor)
		{
			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomFileProcessMethod = (string inputFilePath, string outputFolderPath) =>
				{
					var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

					if (mail == null)
						throw new Exception("Invalid file format");

					AddTextWatermark(inputFilePath, mail, outputFolderPath, watermarkText, fontFamily, watermarkColor);
				}
			};

			return Task.FromResult(processor.Process(folderName, fileName));
		}

		private void AddTextWatermark(string mailFilePath, MapiMessage mail, string outputFolderPath, string watermarkText, string fontFamily, string watermarkColor)
		{
			var html = mail.BodyHtml;

			//var imageBase64 = "";//Convert.ToBase64String(System.IO.File.ReadAllBytes(Path.Combine(AppSettings.WorkingDirectory, imageFoler, imageFile)));
			var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");
			var options = new PngOptions();
			options.Source = new StreamSource(new MemoryStream(), true);

			var width = watermarkText.Length * 17;
			var height = 120;

			using (Image image = Image.Create(options, width, height))
			{
				// Create graphics object to perform draw operations.
				Graphics graphics = new Graphics(image);

				// Create font to draw watermark with.
				Font font = new Font(fontFamily, 20.0f);

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

			var folderPath = Path.Combine(Config.Configuration.OutputDirectory, Guid.NewGuid().ToString());
			var filePath = Path.Combine(folderPath, "Merged.html");

			htmlDocument.Save(filePath);

			var content = System.IO.File.ReadAllText(filePath);

			Directory.Delete(folderPath, true);

			mail.SetBodyContent(content, BodyContentType.Html);

			if (mailFilePath.EndsWith("MSG", StringComparison.OrdinalIgnoreCase))
				mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(mailFilePath) + " Marked.msg"), SaveOptions.DefaultMsgUnicode);
			else
				mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(mailFilePath) + " Marked.eml"), SaveOptions.DefaultEml);
		}

		///<Summary>
		/// ImageWatermark method to add image watermark
		///</Summary>
		[HttpGet]
		[AcceptVerbs("GET", "POST")]
		public Task<Response> ImageWatermark(string fileName, string folderName, string imageFileName, string imageFolderName, string outputType, bool grayScale = false)
		{
			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomFileProcessMethod = (string inputFilePath, string outputFolderPath) =>
				{
					var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

					if (mail == null)
						throw new Exception("Invalid file format");


                    var path = Path.Combine(Config.Configuration.WorkingDirectory, imageFolderName, imageFileName);

                    if (grayScale)
                    {
                        var ms = new MemoryStream();

                        using (var image = Image.Load(path))
                        {
                            var saveOptions = new Aspose.Imaging.ImageOptions.PsdOptions();

                            saveOptions.ColorMode = Imaging.FileFormats.Psd.ColorModes.Grayscale;
                            saveOptions.CompressionMethod = Imaging.FileFormats.Psd.CompressionMethod.Raw;

                            image.Save(ms, saveOptions);
                        }

                        ms.Position = 0;

                        using (var newImage = Image.Load(ms))
                        {
                            ms = new MemoryStream();
                            newImage.Save(ms, new PngOptions());
                        }

						AddImageWatermark(inputFilePath, mail, outputFolderPath, ms.ToArray());
					}
                    else
						AddImageWatermark(inputFilePath, mail, outputFolderPath, System.IO.File.ReadAllBytes(path));
				}
			};

			return Task.FromResult(processor.Process(folderName, fileName));
		}

		private void AddImageWatermark(string inputFilePath, MapiMessage mail, string outputFolderPath, byte[] imageBytes)
		{
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeImagingLicense();
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

			var folderPath = Path.Combine(Config.Configuration.OutputDirectory, Guid.NewGuid().ToString());
			var filePath = Path.Combine(folderPath, "Merged.html");

			htmlDocument.Save(filePath);

			var content = System.IO.File.ReadAllText(filePath);

			Directory.Delete(folderPath, true);

			mail.SetBodyContent(content, BodyContentType.Html);

			if (inputFilePath.Equals("MSG", StringComparison.OrdinalIgnoreCase))
				mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Marked.msg"), SaveOptions.DefaultMsgUnicode);
			else
				mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Marked.eml"), SaveOptions.DefaultEml);
		}

		///<Summary>
		/// RemoveWatermark method to remove watermark
		///</Summary>
		[HttpGet]
		[AcceptVerbs("GET", "POST")]
		public Task<Response> RemoveWatermark(string fileName, string folderName, string userEmail)
		{
			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomFileProcessMethod = (string inputFilePath, string outputFolderPath) =>
				{
					var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

					if (mail == null)
						throw new Exception("Invalid file format");

					RemoveWatermark(inputFilePath, mail, outputFolderPath);
				}
			};

			return Task.FromResult(processor.Process(folderName, fileName));
		}

		private void RemoveWatermark(string inputFilePath, MapiMessage mail, string outputFolderPath)
		{
			var formatInfo = Email.Tools.FileFormatUtil.DetectFileFormat(inputFilePath);
			
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
				element = (HTMLElement)element.Children[0];

				// Process only default tags for watermark
				if (element.NodeName != "TABLE" &&
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

			var folderPath = Path.Combine(Config.Configuration.OutputDirectory, Guid.NewGuid().ToString());
			var filePath = Path.Combine(folderPath, "Merged.html");

			htmlDocument.Save(filePath);

			var content = System.IO.File.ReadAllText(filePath);

			Directory.Delete(folderPath, true);

			mail.SetBodyContent(content, BodyContentType.Html);

			if (formatInfo.FileFormatType == FileFormatType.Msg)
				mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Unmarked.msg"), SaveOptions.DefaultMsgUnicode);
			else
				mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Unmarked.eml"), SaveOptions.DefaultEml);
		}

		private Color GetColor(string watermarkColor)
		{
			Color color;
			if (watermarkColor != "")
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

		private static ImageOptionsBase GetSaveFormat(string outputType)
		{
			switch (outputType)
			{
				case "bmp":
					return new Aspose.Imaging.ImageOptions.BmpOptions();
				case "gif":
					return new Aspose.Imaging.ImageOptions.GifOptions();
				case "jpeg":
					return new Aspose.Imaging.ImageOptions.JpegOptions();
				case "pdf":
					return new Aspose.Imaging.ImageOptions.PdfOptions();
				case "png":
					return new Aspose.Imaging.ImageOptions.PngOptions();
				case "psd":
					return new Aspose.Imaging.ImageOptions.PsdOptions();
				case "tiff":
					return new Aspose.Imaging.ImageOptions.TiffOptions(Imaging.FileFormats.Tiff.Enums.TiffExpectedFormat.Default);
				case "emf":
					return new Aspose.Imaging.ImageOptions.EmfRasterizationOptions();

			}
			return null;
		}
	}
}
