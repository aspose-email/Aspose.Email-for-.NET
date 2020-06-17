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
		public void ConvertMbox(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, string outputType)
		{
			PrepareOutputType(ref outputType);

			switch (outputType)
			{
				case "eml": ConvertMboxToEml(input, shortResourceNameWithExtension, handler); break;
				case "msg": ConvertMboxToMsg(input, shortResourceNameWithExtension, handler); break;
				case "mbox": ReturnSame(input, shortResourceNameWithExtension, handler); break;
				case "pst": ConvertMboxToPst(input, shortResourceNameWithExtension, handler); break;
				case "mht": ConvertMboxToMht(input, shortResourceNameWithExtension, handler); break;
				case "html": ConvertMboxToHtml(input, shortResourceNameWithExtension, handler); break;
				case "svg": ConvertMboxToSvg(input, shortResourceNameWithExtension, handler); break;
				case "tiff": ConvertMboxToTiff(input, shortResourceNameWithExtension, handler); break;
				case "jpg": ConvertMboxToJpg(input, shortResourceNameWithExtension, handler); break;
				case "bmp": ConvertMboxToBmp(input, shortResourceNameWithExtension, handler); break;
				case "png": ConvertMboxToPng(input, shortResourceNameWithExtension, handler); break;
				case "pdf": ConvertMboxToPdf(input, shortResourceNameWithExtension, handler); break;
				case "doc": ConvertMboxToDoc(input, shortResourceNameWithExtension, handler); break;
				case "ppt": ConvertMboxToPpt(input, shortResourceNameWithExtension, handler); break;
				case "rtf": ConvertMboxToRtf(input, shortResourceNameWithExtension, handler); break;
				case "docx": ConvertMboxToDocx(input, shortResourceNameWithExtension, handler); break;
				case "docm": ConvertMboxToDocm(input, shortResourceNameWithExtension, handler); break;
				case "dotx": ConvertMboxToDotx(input, shortResourceNameWithExtension, handler); break;
				case "dotm": ConvertMboxToDotm(input, shortResourceNameWithExtension, handler); break;
				case "odt": ConvertMboxToOdt(input, shortResourceNameWithExtension, handler); break;
				case "ott": ConvertMboxToOtt(input, shortResourceNameWithExtension, handler); break;
				case "epub": ConvertMboxToEpub(input, shortResourceNameWithExtension, handler); break;
				case "txt": ConvertMboxToTxt(input, shortResourceNameWithExtension, handler); break;
				case "emf": ConvertMboxToEmf(input, shortResourceNameWithExtension, handler); break;
				case "xps": ConvertMboxToXps(input, shortResourceNameWithExtension, handler); break;
				case "pcl": ConvertMboxToPcl(input, shortResourceNameWithExtension, handler); break;
				case "ps": ConvertMboxToPs(input, shortResourceNameWithExtension, handler); break;
				case "mhtml": ConvertMboxToMhtml(input, shortResourceNameWithExtension, handler); break;
				default:
					throw new NotSupportedException($"Output type not supported {outputType.ToUpperInvariant()}");
			}
		}

		public void ConvertMboxToTiff(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveMboxAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Tiff, ".tiff");
		}

		public void ConvertMboxToJpg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveMboxAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Jpeg, ".jpg");
		}

		public void ConvertMboxToBmp(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveMboxAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Bmp, ".bmp");
		}

		public void ConvertMboxToPng(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveMboxAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Png, ".png");
		}

		public void ConvertMboxToSvg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveMboxAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Svg, ".svg");
		}

		public void ConvertMboxToHtml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveMboxAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Html, ".html");
		}

		public void ConvertMboxToMht(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
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

			using (var reader = new MboxrdStorageReader(input, false))
			{
				for (int i = 0; i < reader.GetTotalItemsCount(); i++)
				{
					var message = reader.ReadNextMessage();

					using (var output = handler.CreateOutputStream("Message" + i + ".mht"))
						message.Save(output, mhtSaveOptions);
				}
			}
		}

		///<Summary>
		/// ConvertMboxToPst method to convert mbox to pst file
		///</Summary>
		public void ConvertMboxToPst(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var personalStorage = PersonalStorage.Create(handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".pst")), FileFormatVersion.Unicode))
			{
				var inbox = personalStorage.RootFolder.AddSubFolder("Inbox");

				using (MboxrdStorageReader reader = new MboxrdStorageReader(input, false))
				{
					MailMessage message;

					while ((message = reader.ReadNextMessage()) != null)
					{
						using (var mapi = MapiMessage.FromMailMessage(message, MapiConversionOptions.UnicodeFormat))
							inbox.AddMessage(mapi);
					}
				}
			}
		}

		///<Summary>
		/// ConvertMboxToEml method to convert mbox to eml file
		///</Summary>
		public void ConvertMboxToEml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var reader = new MboxrdStorageReader(input, false))
			{
				for (int i = 0; i < reader.GetTotalItemsCount(); i++)
				{
					using (var message = reader.ReadNextMessage())
					{
						using (var output = handler.CreateOutputStream("Message" + i + ".eml"))
							message.Save(output, SaveOptions.DefaultEml);
					}
				}
			}
		}

		///<Summary>
		/// ConvertMboxToMsg method to convert mbox to msg file
		///</Summary>
		public void ConvertMboxToMsg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var reader = new MboxrdStorageReader(input, false))
			{
				for (int i = 0; i < reader.GetTotalItemsCount(); i++)
				{
					using (var message = reader.ReadNextMessage())
					{
						using (var output = handler.CreateOutputStream("Message" + i + ".msg"))
							message.Save(output, SaveOptions.DefaultMsgUnicode);
					}
				}
			}
		}

		public void ConvertMboxToPdf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Pdf);
		}

		public void ConvertMboxToDoc(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Doc);
		}

		public void ConvertMboxToPpt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var reader = new MboxrdStorageReader(input, false))
			{
				using (var presentation = new Presentation())
				{
					var firstSlide = presentation.Slides[0];
					var convertOptions = new MapiConversionOptions();
					var renameCallback = new ImageSavingCallback(handler);

					for (int i = 0; i < reader.GetTotalItemsCount(); i++)
					{
						var message = reader.ReadNextMessage();
						var newSlide = presentation.Slides.AddClone(firstSlide);
						AddMessageInSlide(presentation, newSlide, MapiMessage.FromMailMessage(message, convertOptions), renameCallback);
					}

					presentation.Slides.Remove(firstSlide);

					using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".ppt")))
						presentation.Save(output, Aspose.Slides.Export.SaveFormat.Ppt);
				}
			}
		}

		public void ConvertMboxToRtf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Rtf);
		}

		public void ConvertMboxToDocx(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Docx);
		}

		public void ConvertMboxToDocm(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Docm);
		}

		public void ConvertMboxToDotx(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Dotx);
		}

		public void ConvertMboxToDotm(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Dotm);
		}

		public void ConvertMboxToOdt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Odt);
		}

		public void ConvertMboxToOtt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Ott);
		}

		public void ConvertMboxToEpub(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Epub);
		}

		public void ConvertMboxToTxt(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Text);
		}

		public void ConvertMboxToEmf(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Emf);
		}

		public void ConvertMboxToXps(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Xps);
		}

		public void ConvertMboxToPcl(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Pcl);
		}

		public void ConvertMboxToPs(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Ps);
		}

		public void ConvertMboxToMhtml(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			SaveOstOrPstAsDocument(input, shortResourceNameWithExtension, handler, Words.SaveFormat.Mhtml);
		}

		void SaveMboxAsDocument(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, Words.SaveFormat format, string saveExtension)
		{
			using (var reader = new MboxrdStorageReader(input, false))
			{
				var msgStream = new MemoryStream();

				for (int i = 0; i < reader.GetTotalItemsCount(); i++)
				{
					var message = reader.ReadNextMessage();
					message.Save(msgStream, SaveOptions.DefaultMhtml);
				}

				msgStream.Position = 0;

				SaveDocumentStreamToFolder(msgStream, shortResourceNameWithExtension, handler, format);
			}
		}
	}
}
