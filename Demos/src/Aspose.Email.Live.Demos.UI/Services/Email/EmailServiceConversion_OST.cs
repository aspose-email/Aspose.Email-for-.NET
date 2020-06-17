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
		public void ConvertOst(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, string outputType)
		{
			PrepareOutputType(ref outputType);

			switch (outputType)
			{
				case "eml": ConvertOstToEml(input, shortResourceNameWithExtension, handler); break;
				case "msg": ConvertOstToMsg(input, shortResourceNameWithExtension, handler); break;
				case "mbox": ConvertOstToMbox(input, shortResourceNameWithExtension, handler); break;
				case "ost": ReturnSame(input, shortResourceNameWithExtension, handler); break;
				case "pst": ConvertOSTToPST(input, shortResourceNameWithExtension, handler); break;
				case "mht": ConvertOstToMht(input, shortResourceNameWithExtension, handler); break;
				case "html": ConvertOstToHtml(input, shortResourceNameWithExtension, handler); break;
				case "svg": ConvertOstToSvg(input, shortResourceNameWithExtension, handler); break;
				case "tiff": ConvertOstToTiff(input, shortResourceNameWithExtension, handler); break;
				case "jpg": ConvertOstToJpg(input, shortResourceNameWithExtension, handler); break;
				case "bmp": ConvertOstToBmp(input, shortResourceNameWithExtension, handler); break;
				case "png": ConvertOstToPng(input, shortResourceNameWithExtension, handler); break;
				case "pdf": ConvertOstToPdf(input, shortResourceNameWithExtension, handler); break;
				case "doc": ConvertOstToDoc(input, shortResourceNameWithExtension, handler); break;
				case "ppt": ConvertOstToPpt(input, shortResourceNameWithExtension, handler); break;
				case "rtf": ConvertOstToRtf(input, shortResourceNameWithExtension, handler); break;
				case "docx": ConvertOstToDocx(input, shortResourceNameWithExtension, handler); break;
				case "docm": ConvertOstToDocm(input, shortResourceNameWithExtension, handler); break;
				case "dotx": ConvertOstToDotx(input, shortResourceNameWithExtension, handler); break;
				case "dotm": ConvertOstToDotm(input, shortResourceNameWithExtension, handler); break;
				case "odt": ConvertOstToOdt(input, shortResourceNameWithExtension, handler); break;
				case "ott": ConvertOstToOtt(input, shortResourceNameWithExtension, handler); break;
				case "epub": ConvertOstToEpub(input, shortResourceNameWithExtension, handler); break;
				case "txt": ConvertOstToTxt(input, shortResourceNameWithExtension, handler); break;
				case "emf": ConvertOstToEmf(input, shortResourceNameWithExtension, handler); break;
				case "xps": ConvertOstToXps(input, shortResourceNameWithExtension, handler); break;
				case "pcl": ConvertOstToPcl(input, shortResourceNameWithExtension, handler); break;
				case "ps": ConvertOstToPs(input, shortResourceNameWithExtension, handler); break;
				case "mhtml": ConvertOstToMhtml(input, shortResourceNameWithExtension, handler); break;
				default:
					throw new NotSupportedException($"Output type not supported {outputType.ToUpperInvariant()}");
			}
		}


		public void ConvertOstToPng(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Png);
		}

		public void ConvertOstToBmp(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Bmp);
		}

		public void ConvertOstToJpg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Jpeg);
		}

		public void ConvertOstToTiff(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Tiff);
		}

		public void ConvertOstToSvg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Svg);
		}

		public void ConvertOstToHtml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Html);
		}

		public void ConvertOstToMbox(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var personalStorage = PersonalStorage.FromStream(input))
			{
				var options = new MailConversionOptions();

				using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".mbox")))
				{
					using (var writer = new MboxrdStorageWriter(output, false))
					{
						HandleFolderAndSubfolders(mapiMessage =>
						{
							using (var msg = mapiMessage.ToMailMessage(options))
								writer.WriteMessage(msg);
						}, personalStorage.RootFolder, options);
					}
				}
			}
		}

		public void ConvertOstToMht(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
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

			using (var personalStorage = PersonalStorage.FromStream(input))
			{
				int i = 0;
				HandleFolderAndSubfolders(mapiMessage =>
				{
					using (var output = handler.CreateOutputStream("Message" + i++ + ".mht"))
						mapiMessage.Save(output, mhtSaveOptions);
				}, personalStorage.RootFolder, new MailConversionOptions());
			}
		}

		public void ConvertOstToMsg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var personalStorage = PersonalStorage.FromStream(input))
			{
				int i = 0;
				HandleFolderAndSubfolders(mapiMessage =>
				{
					using (var output = handler.CreateOutputStream("Message" + i++ + ".msg"))
						mapiMessage.Save(output, SaveOptions.DefaultMsgUnicode);
				}, personalStorage.RootFolder, new MailConversionOptions());
			}
		}

		public void ConvertOstToEml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var personalStorage = PersonalStorage.FromStream(input))
			{
				int i = 0;
				HandleFolderAndSubfolders(mapiMessage =>
				{
					using (var output = handler.CreateOutputStream("Message" + i++ + ".eml"))
						mapiMessage.Save(output, SaveOptions.DefaultEml);
				}, personalStorage.RootFolder, new MailConversionOptions());
			}
		}

		public void ConvertOSTToPST(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			// 'SaveAs' to stream currently not supported by Aspose.Email.
			// That is why convert opened ost in memory and save memoryStream to output

			using (var memoryStream = new MemoryStream())
			{
				input.CopyTo(memoryStream);

				using (var personalStorage = PersonalStorage.FromStream(memoryStream))
					personalStorage.ConvertTo(FileFormat.Pst);

				using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".pst")))
					memoryStream.CopyTo(output);
			}
		}

		public void ConvertOstToPdf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Pdf);
		}

		public void ConvertOstToDoc(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Doc);
		}

		public void ConvertOstToPpt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var personalStorage = PersonalStorage.FromStream(input))
			{
				using (var presentation = new Presentation())
				{
					var firstSlide = presentation.Slides[0];
					var renameCallback = new ImageSavingCallback(handler);

					HandleFolderAndSubfolders(mapiMessage =>
					{
						var newSlide = presentation.Slides.AddClone(firstSlide);
						AddMessageInSlide(presentation, newSlide, mapiMessage, renameCallback);
					}, personalStorage.RootFolder, new MailConversionOptions());

					presentation.Slides.Remove(firstSlide);

					using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".ppt")))
						presentation.Save(output, Aspose.Slides.Export.SaveFormat.Ppt);
				}
			}
		}

		public void ConvertOstToRtf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Rtf);
		}

		public void ConvertOstToDocx(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Docx);
		}

		public void ConvertOstToDocm(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Docm);
		}

		public void ConvertOstToDotx(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Dotx);
		}

		public void ConvertOstToDotm(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Dotm);
		}

		public void ConvertOstToOdt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Odt);
		}

		public void ConvertOstToOtt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Ott);
		}

		public void ConvertOstToEpub(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Epub);
		}

		public void ConvertOstToTxt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Text);
		}

		public void ConvertOstToEmf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Emf);
		}

		public void ConvertOstToXps(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Xps);
		}

		public void ConvertOstToPcl(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Pcl);
		}

		public void ConvertOstToPs(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Ps);
		}

		public void ConvertOstToMhtml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Mhtml);
		}
	}
}
