using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using Aspose.Slides;
using System;
using System.IO;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	public partial class EmailService
	{
		public void ConvertEmlOrMsg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, string outputType)
		{
			PrepareOutputType(ref outputType);

			switch (outputType)
			{
				case "eml": ConvertMIMEMessageToEML(input, shortResourceNameWithExtension, handler); break;
				case "msg": ConvertEMLToMSG(input, shortResourceNameWithExtension, handler); break;
				case "mbox": ConvertMailToMbox(input, shortResourceNameWithExtension, handler); break;
				case "pst": ConvertMailToPst(input, shortResourceNameWithExtension, handler); break;
				case "mht": ConvertEMLToMHT(input, shortResourceNameWithExtension, handler); break;
				case "html": ConvertEmailToHtml(input, shortResourceNameWithExtension, handler); break;
				case "svg":
				case "tiff": ConvertEmailToSingleImage(input, shortResourceNameWithExtension, handler, outputType); break;
				case "jpg":
				case "bmp":
				case "png": ConvertEmailToImages(input, shortResourceNameWithExtension, handler, outputType); break;
				case "pdf": ConvertEmailToPdf(input, shortResourceNameWithExtension, handler); break;
				case "doc": ConvertEmailToDoc(input, shortResourceNameWithExtension, handler); break;
				case "ppt": ConvertEmailToPpt(input, shortResourceNameWithExtension, handler); break;
				case "rtf": ConvertEmailToRtf(input, shortResourceNameWithExtension, handler); break;
				case "docx": ConvertEmailToDocx(input, shortResourceNameWithExtension, handler); break;
				case "docm": ConvertEmailToDocm(input, shortResourceNameWithExtension, handler); break;
				case "dotx": ConvertEmailToDotx(input, shortResourceNameWithExtension, handler); break;
				case "dotm": ConvertEmailToDotm(input, shortResourceNameWithExtension, handler); break;
				case "odt": ConvertEmailToOdt(input, shortResourceNameWithExtension, handler); break;
				case "ott": ConvertEmailToOtt(input, shortResourceNameWithExtension, handler); break;
				case "epub": ConvertEmailToEpub(input, shortResourceNameWithExtension, handler); break;
				case "txt": ConvertEmailToTxt(input, shortResourceNameWithExtension, handler); break;
				case "emf": ConvertEmailToEmf(input, shortResourceNameWithExtension, handler); break;
				case "xps": ConvertEmailToXps(input, shortResourceNameWithExtension, handler); break;
				case "pcl": ConvertEmailToPcl(input, shortResourceNameWithExtension, handler); break;
				case "ps": ConvertEmailToPs(input, shortResourceNameWithExtension, handler); break;
				case "mhtml": ConvertEmailToMhtml(input, shortResourceNameWithExtension, handler); break;
				default:
					throw new NotSupportedException($"Output type not supported {outputType?.ToUpperInvariant()}");
			}
		}

		///<Summary>
		/// Convert eml message to msg 
		///</Summary>
		public void ConvertEMLToMSG(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			// Load mail message
			using (var message = MailMessage.Load(input, new EmlLoadOptions()))
			{
				// Save as MSG
				using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".msg")))
					message.Save(output, SaveOptions.DefaultMsgUnicode);
			}
		}

		///<Summary>
		/// Convert mime message to eml 
		///</Summary>
		public void ConvertMIMEMessageToEML(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			// Load mail message
			using (var message = MailMessage.Load(input))
			{
				// Save as EML
				using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".eml")))
					message.Save(output, SaveOptions.DefaultEml);
			}
		}

		/// <summary>
		/// Convert mime message to mbox 
		/// </summary>
		public void ConvertMailToMbox(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var message = MailMessage.Load(input))
			{
				using (var outputStream = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".mbox")))
				{
					using (var writer = new MboxrdStorageWriter(outputStream, false))
					{
						writer.WriteMessage(message);
					}
				}
			}
		}

		/// <summary>
		/// Convert mime message to pst 
		/// </summary>
		public void ConvertMailToPst(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var message = MailMessage.Load(input))
			{
				using (var outputStream = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".pst")))
				{
					using (var personalStorage = PersonalStorage.Create(outputStream, FileFormatVersion.Unicode))
					{
						var inbox = personalStorage.RootFolder.AddSubFolder("Inbox");

						inbox.AddMessage(MapiMessage.FromMailMessage(message, MapiConversionOptions.UnicodeFormat));
					}
				}
			}
		}

		///<Summary>
		/// Convert eml to mht
		///</Summary>
		public void ConvertEMLToMHT(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			// Load mail message
			using (var mail = MailMessage.Load(input, new EmlLoadOptions()))
			{
				// Save as mht with header
				var mhtSaveOptions = new MhtSaveOptions
				{
					//Specify formatting options required
					//Here we are specifying to write header informations to output without writing extra print header
					//and the output headers should display as the original headers in message
					MhtFormatOptions = MhtFormatOptions.WriteHeader | MhtFormatOptions.HideExtraPrintHeader | MhtFormatOptions.DisplayAsOutlook,
					// Check the body encoding for validity. 
					CheckBodyContentEncoding = true
				};

				using (var outputStream = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".mht")))
					mail.Save(outputStream, mhtSaveOptions);
			}
		}

		///<Summary>
		/// ConvertEmailToHtml method to email to html
		///</Summary>
		public void ConvertEmailToHtml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var mail = MailMessage.Load(input, new EmlLoadOptions()))
			{
				var htmlSaveOptions = new HtmlSaveOptions
				{
					HtmlFormatOptions = HtmlFormatOptions.WriteHeader | HtmlFormatOptions.DisplayAsOutlook,
					CheckBodyContentEncoding = true
				};

				using (var outputStream = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".html")))
					mail.Save(outputStream, htmlSaveOptions);
			}
		}

		///<Summary>
		/// ConvertEmailToImages method to convert email to messages
		///</Summary>
		public void ConvertEmailToImages(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, string outputType)
		{
			if (outputType != null)
			{
				if (outputType.Equals("bmp", StringComparison.OrdinalIgnoreCase) ||
					outputType.Equals("jpg", StringComparison.OrdinalIgnoreCase) ||
					outputType.Equals("png", StringComparison.OrdinalIgnoreCase))
				{
					var format = Aspose.Words.SaveFormat.Bmp;

					if (outputType.Equals("jpg", StringComparison.OrdinalIgnoreCase))
					{
						format = Aspose.Words.SaveFormat.Jpeg;
					}
					else if (outputType.Equals("png", StringComparison.OrdinalIgnoreCase))
					{
						format = Aspose.Words.SaveFormat.Png;
					}

					using (var mail = MailMessage.Load(input))
					{
						var msgStream = new MemoryStream();
						mail.Save(msgStream, SaveOptions.DefaultMhtml);
						msgStream.Position = 0;

						var options = new Aspose.Words.Saving.ImageSaveOptions(format);
						options.PageCount = 1;

						var outfileName = Path.GetFileNameWithoutExtension(shortResourceNameWithExtension) + "_{0}";

						var doc = new Aspose.Words.Document(msgStream);
						var pageCount = doc.PageCount;

						for (int i = 0; i < doc.PageCount; i++)
						{
							if (pageCount > 1)
							{
								options.PageIndex = i;

								using (var outputStream = handler.CreateOutputStream(string.Format(outfileName, (i + 1) + "." + outputType)))
									doc.Save(outputStream, options);
							}
							else
							{
								options.PageIndex = i;

								using (var outputStream = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, "." + outputType)))
									doc.Save(outputStream, options);
							}

						}
					}
				}
			}

			throw new NotSupportedException($"Output type not supported {outputType?.ToUpperInvariant()}");
		}

		///<Summary>
		/// ConvertEmailToSingleImage method to convert email to single image
		///</Summary>
		public void ConvertEmailToSingleImage(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, string outputType)
		{
			if (outputType != null)
			{
				if (outputType.Equals("tiff") || outputType.Equals("svg"))
				{
					var format = Aspose.Words.SaveFormat.Tiff;

					if (outputType.Equals("svg"))
					{
						format = Aspose.Words.SaveFormat.Svg;
					}

					SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, format);
				}
			}

			throw new NotSupportedException($"Output type not supported {outputType?.ToUpperInvariant()}");
		}

		public void ConvertEmailToPdf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Pdf);
		}

		public void ConvertEmailToDoc(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Doc);
		}

		public void ConvertEmailToPpt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var presentation = new Presentation())
			{
				ISlide slide = presentation.Slides[0];

				using (var msg = MapiHelper.GetMailMessageFromStream(input))
				{
					AddMessageInSlide(presentation, slide, msg, new ImageSavingCallback(handler));

					using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".ppt")))
						presentation.Save(output, Aspose.Slides.Export.SaveFormat.Ppt);
				}
			}
		}

		void AddMessageInSlide(IPresentation presentation, ISlide slide, MapiMessage msg, ImageSavingCallback renameCallback)
		{
			//Adding the AutoShape to accomodate the HTML content
			IAutoShape ashape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, presentation.SlideSize.Size.Width - 20, presentation.SlideSize.Size.Height - 10);

			ashape.FillFormat.FillType = FillType.NoFill;

			//Adding text frame to the shape
			ashape.AddTextFrame("");

			//Clearing all paragraphs in added text frame
			ashape.TextFrame.Paragraphs.Clear();


			var msgStream = new MemoryStream();
			msg.Save(msgStream, SaveOptions.DefaultMhtml);
			msgStream.Position = 0;

			var options = new Aspose.Words.Saving.HtmlSaveOptions();
			options.DocumentSplitCriteria = Words.Saving.DocumentSplitCriteria.SectionBreak;

			var doc = new Aspose.Words.Document(msgStream);
			var docStream = new MemoryStream();

			// If we convert a document that contains images into html, we will end up with one html file which links to several images
			// Each image will be in the form of a file in the local file system
			// There is also a callback that can customize the name and file system location of each image
			options.ImageSavingCallback = renameCallback;

			doc.Save(docStream, options);
			docStream.Position = 0;

			var html = new StreamReader(docStream).ReadToEnd();

			ashape.TextFrame.Paragraphs.AddFromHtml(html);
		}

		public void ConvertEmailToRtf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var mapi = MapiHelper.GetMailMessageFromStream(input))
			{
				using (var outputStream = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".rtf")))
				using (var writer = new StreamWriter(outputStream))
					writer.Write(mapi.BodyRtf);
			}
		}

		public void ConvertEmailToDocx(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Docx);
		}

		public void ConvertEmailToDotx(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Dotx);
		}

		public void ConvertEmailToOdt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Odt);
		}

		public void ConvertEmailToOtt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Ott);
		}

		public void ConvertEmailToEpub(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Epub);
		}

		public void ConvertEmailToTxt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var mapi = MapiHelper.GetMailMessageFromStream(input))
			{
				using (var outputStream = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".txt")))
				using (var writer = new StreamWriter(outputStream))
					writer.Write(mapi.Body);
			}
		}

		public void ConvertEmailToPs(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Ps);
		}

		public void ConvertEmailToPcl(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Pcl);
		}

		public void ConvertEmailToXps(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Xps);
		}

		public void ConvertEmailToEmf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Emf);
		}

		public void ConvertEmailToDocm(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Docm);
		}

		public void ConvertEmailToDotm(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Dotm);
		}

		public void ConvertEmailToMhtml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveByWordsWithFormat(input, shortResourceNameWithExtension, handler, Aspose.Words.SaveFormat.Mhtml);
		}
	}
}
