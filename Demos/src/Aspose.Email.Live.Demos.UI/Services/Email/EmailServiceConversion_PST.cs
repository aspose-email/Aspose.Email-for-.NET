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
		public void ConvertPst(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, string outputType)
		{
			PrepareOutputType(ref outputType);

			switch (outputType)
			{
				case "eml": ConvertPstToEml(input, shortResourceNameWithExtension, handler); break;
				case "msg": ConvertPstToMsg(input, shortResourceNameWithExtension, handler); break;
				case "mbox": ConvertPstToMbox(input, shortResourceNameWithExtension, handler); break;
				case "pst": ReturnSame(input, shortResourceNameWithExtension, handler); break;
				case "mht": ConvertPstToMht(input, shortResourceNameWithExtension, handler); break;
				case "html": ConvertPstToHtml(input, shortResourceNameWithExtension, handler); break;
				case "svg": ConvertPstToSvg(input, shortResourceNameWithExtension, handler); break;
				case "tiff": ConvertPstToTiff(input, shortResourceNameWithExtension, handler); break;
				case "jpg": ConvertPstToJpg(input, shortResourceNameWithExtension, handler); break;
				case "bmp": ConvertPstToBmp(input, shortResourceNameWithExtension, handler); break;
				case "png": ConvertPstToPng(input, shortResourceNameWithExtension, handler); break;
				case "pdf": ConvertPstToPdf(input, shortResourceNameWithExtension, handler); break;
				case "doc": ConvertPstToDoc(input, shortResourceNameWithExtension, handler); break;
				case "ppt": ConvertPstToPpt(input, shortResourceNameWithExtension, handler); break;
				case "rtf": ConvertPstToRtf(input, shortResourceNameWithExtension, handler); break;
				case "docx": ConvertPstToDocx(input, shortResourceNameWithExtension, handler); break;
				case "docm": ConvertPstToDocm(input, shortResourceNameWithExtension, handler); break;
				case "dotx": ConvertPstToDotx(input, shortResourceNameWithExtension, handler); break;
				case "dotm": ConvertPstToDotm(input, shortResourceNameWithExtension, handler); break;
				case "odt": ConvertPstToOdt(input, shortResourceNameWithExtension, handler); break;
				case "ott": ConvertPstToOtt(input, shortResourceNameWithExtension, handler); break;
				case "epub": ConvertPstToEpub(input, shortResourceNameWithExtension, handler); break;
				case "txt": ConvertPstToTxt(input, shortResourceNameWithExtension, handler); break;
				case "emf": ConvertPstToEmf(input, shortResourceNameWithExtension, handler); break;
				case "xps": ConvertPstToXps(input, shortResourceNameWithExtension, handler); break;
				case "pcl": ConvertPstToPcl(input, shortResourceNameWithExtension, handler); break;
				case "ps": ConvertPstToPs(input, shortResourceNameWithExtension, handler); break;
				case "mhtml": ConvertPstToMhtml(input, shortResourceNameWithExtension, handler); break;
				default:
					throw new NotSupportedException($"Output type not supported {outputType.ToUpperInvariant()}");
			}
		}


		public void ConvertPstToMsg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
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

		public void ConvertPstToMht(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
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

		public void ConvertPstToHtml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Html);
		}

		public void ConvertPstToSvg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Svg);
		}

		public void ConvertPstToTiff(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Tiff);
		}

		public void ConvertPstToJpg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Jpeg);
		}

		public void ConvertPstToBmp(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Bmp);
		}

		public void ConvertPstToPng(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Png);
		}

		public void ConvertPstToEml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
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

		///<Summary>
		/// ConvertPstToMbox method to convert pst to mbox file
		///</Summary>
		public void ConvertPstToMbox(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			var options = new MailConversionOptions();

			using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".mbox")))
			{
				using (var writer = new MboxrdStorageWriter(output, false))
				{
					using (var pst = PersonalStorage.FromStream(input))
					{
						HandleFolderAndSubfolders(mapiMessage =>
						{
							var msg = mapiMessage.ToMailMessage(options);
							writer.WriteMessage(msg);
						}, pst.RootFolder, options);
					}
				}
			}
		}

		public void ConvertPstToPdf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Pdf);
		}

		public void ConvertPstToDoc(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Doc);
		}

		public void ConvertPstToPpt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
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

		public void ConvertPstToRtf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Rtf);
		}

		public void ConvertPstToDocx(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Docx);
		}

		public void ConvertPstToDocm(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Docm);
		}

		public void ConvertPstToDotx(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Dotx);
		}

		public void ConvertPstToDotm(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Dotm);
		}

		public void ConvertPstToOdt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Odt);
		}

		public void ConvertPstToOtt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Ott);
		}

		public void ConvertPstToEpub(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Epub);
		}

		public void ConvertPstToTxt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Text);
		}

		public void ConvertPstToEmf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Emf);
		}

		public void ConvertPstToXps(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Xps);
		}

		public void ConvertPstToPcl(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Pcl);
		}

		public void ConvertPstToPs(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Ps);
		}

		public void ConvertPstToMhtml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Mhtml);
		}
	}
}
