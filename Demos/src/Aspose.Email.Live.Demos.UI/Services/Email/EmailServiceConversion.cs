using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System;
using System.IO;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	public partial class EmailService
    {
		public void Convert(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, string outputType)
		{
			var ext = Path.GetExtension(shortResourceNameWithExtension);

			/* All conversions table:
			 * 
			 * INPUTS   OUTPUTS
			 * 
			 * eml		eml
			 * msg		msg
			 * mbox		mbox
			 * ost		pst
			 * pst		mht
			 *			html
			 *			svg
			 *			tiff
			 *			jpg
			 *			bmp
			 *			png
			 *			pdf
			 *			doc
			 *			ppt
			 *			rtf
			 *			docx
			 *			docm
			 *			dotx
			 *			dotm
			 *			odt
			 *			ott
			 *			epub
		     *			txt
			 *			emf
			 *			xps
			 *			pcl
			 *			ps
			 *			mhtml
			 */

			switch (ext.ToLowerInvariant())
			{
				case ".eml":
				case ".msg": ConvertEmlOrMsg(input, shortResourceNameWithExtension, handler, outputType); break;
				case ".mbox": ConvertMbox(input, shortResourceNameWithExtension, handler, outputType); break;
				case ".ost": ConvertOst(input, shortResourceNameWithExtension, handler, outputType); break;
				case ".pst": ConvertPst(input, shortResourceNameWithExtension, handler, outputType); break;
				default:
					throw new NotSupportedException($"Input type not supported {ext?.ToUpperInvariant()}");
			}
		}

		void ReturnSame(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var output = handler.CreateOutputStream(shortResourceNameWithExtension))
				input.CopyTo(output);
		}

		void SaveOstOrPstAsDocument(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, Words.SaveFormat format)
		{
			using (var personalStorage = PersonalStorage.FromStream(input))
			{
				var msgStream = new MemoryStream();

				HandleFolderAndSubfolders(mapiMessage =>
				{
					mapiMessage.Save(msgStream, SaveOptions.DefaultMhtml);
				}, personalStorage.RootFolder, new MailConversionOptions());

				msgStream.Position = 0;
				SaveDocumentStreamToFolder(msgStream, shortResourceNameWithExtension, handler, format);
			}
		}

		void HandleFolderAndSubfolders(Action<MapiMessage> handler, FolderInfo folderInfo, MailConversionOptions options)
		{
			foreach (MapiMessage mapiMessage in folderInfo.EnumerateMapiMessages())
			{
				handler(mapiMessage);
			}

			if (folderInfo.HasSubFolders == true)
			{
				foreach (FolderInfo subfolderInfo in folderInfo.GetSubFolders())
				{
					HandleFolderAndSubfolders(handler, subfolderInfo, options);
				}
			}
		}

		void SaveDocumentStreamToFolder(Stream stream, string shortResourceNameWithExtension, IOutputHandler handler, Words.SaveFormat format)
		{
			var document = new Words.Document(stream);
			var shortResourceName = Path.GetFileNameWithoutExtension(shortResourceNameWithExtension);
			var formatExt = "." + format.ToString().ToLowerInvariant();

			switch (format)
			{
				//case SaveFormat.Svg:
				//case SaveFormat.Tiff:
				case Words.SaveFormat.Png:
				case Words.SaveFormat.Bmp:
				case Words.SaveFormat.Emf:
				case Words.SaveFormat.Jpeg:
				case Words.SaveFormat.Gif:
					{
						var pageCount = document.PageCount;

						if (pageCount == 1)
						{
							using (var output = handler.CreateOutputStream(shortResourceName + formatExt))
								document.Save(output, format);
							break;
						}

						var options = new Aspose.Words.Saving.ImageSaveOptions(format);
						options.PageCount = 1;

						for (int i = 0; i < document.PageCount; i++)
						{
							options.PageIndex = i;

							using (var output = handler.CreateOutputStream(shortResourceName + i + formatExt))
								document.Save(output, format);
						}
						break;
					}
				default:
					{
						using (var output = handler.CreateOutputStream(shortResourceName + formatExt))
							document.Save(output, format);
						break;
					}
			}
		}

		void SaveByWordsWithFormat(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, Aspose.Words.SaveFormat format)
		{
			using (var msg = MailMessage.Load(input))
			{
				var msgStream = new MemoryStream();
				msg.Save(msgStream, Aspose.Email.SaveOptions.DefaultMhtml);
				msgStream.Position = 0;

				var msgDocument = new Aspose.Words.Document(msgStream);

				using (var outputStream = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, "." + format.ToString().ToLowerInvariant())))
					msgDocument.Save(outputStream, format);
			}
		}

		void PrepareOutputType(ref string outputType)
		{
			if (outputType == null)
				throw new ArgumentNullException(nameof(outputType));

			if (outputType.StartsWith(".", StringComparison.OrdinalIgnoreCase))
				outputType = outputType.Substring(1);

			outputType = outputType.ToLowerInvariant();
		}

		/// <summary>
		/// Renames saved images that are produced when an HTML document is saved.
		/// </summary>
		class ImageSavingCallback : Words.Saving.IImageSavingCallback
		{
			readonly IOutputHandler _handler;
			int _counter;
			public ImageSavingCallback(IOutputHandler handler)
			{
				_handler = handler;
			}

			public void ImageSaving(Words.Saving.ImageSavingArgs args)
			{
				var name = $"Image{_counter++}{Path.GetExtension(args.ImageFileName)}";
				args.ImageFileName = name;
				args.ImageStream = _handler.CreateOutputStream(name); //new FileStream(Path.Combine(_outputDirectory, fileName), FileMode.Create);
			}
		}
	}
}
