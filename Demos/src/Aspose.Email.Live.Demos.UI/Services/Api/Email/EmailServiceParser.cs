using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using Aspose.Words;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Linq;


namespace Aspose.Email.Live.Demos.UI.Services.Email
{
    public partial class EmailService
	{
		public void Parse(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			var ext = Path.GetExtension(shortResourceNameWithExtension);

			switch (ext.ToLowerInvariant())
			{
				case ".eml":
				case ".msg": ParseEmlOrMsg(input, shortResourceNameWithExtension, handler); break;
				case ".mbox": ParseMbox(input, shortResourceNameWithExtension, handler); break;
				case ".ost":
				case ".pst": ParseOstOrPst(input, shortResourceNameWithExtension, handler); break;
				default:
					throw new NotSupportedException($"Input type not supported {ext?.ToUpperInvariant()}");
			}
		}

		public FileDataResult ParseFileData(Stream input)
		{
			var result = new FileDataResult();

			using (var mail = MapiHelper.GetMapiMessageFromStream(input))
			{
				result.Html = mail.BodyHtml;

				// May there are only 3 attachments if the Aspose.Email license has expired
				// Renew the license to fix it

				if (mail.Attachments != null)
					result.Attachments = mail.Attachments.Select(x => new FileAttachment()
					{
						Name = x.FileName,
						Size = x.BinaryData?.Length.ToString() ?? "---",
						Source = x.BinaryData,
						Type = GetFileTypeByExtension(Path.GetExtension(x.FileName)),
						Extension = Path.GetExtension(x.FileName)?.ToLowerInvariant()
					}).ToArray();
			}

			return result;
		}

		private string GetFileTypeByExtension(string extension)
		{
			switch (extension?.ToLowerInvariant())
			{
				case ".eml": return "Mail Message";
				case ".msg": return "Mail Message";
				case ".mbox": return "Message Box";
				case ".ost": return "Open Storage";
				case ".pst": return "Personal Storage";
				case ".pdf": return "Pdf";
				case ".pptx": return "Presentation";
				case ".ppt": return "Presentation";
				case ".odt": return "Document";
				case ".ott": return "Document";
				case ".dotx": return "Document";
				case ".dotm": return "Document";
				case ".docx": return "Document";
				case ".docm": return "Document";
				case ".doc": return "Document";
				case ".rtf": return "Document";
				case ".txt": return "Text";
				case ".jpg": return "Image";
				case ".jpeg": return "Image";
				case ".png": return "Image";
				case ".tiff": return "Image";
				case ".bmp": return "Image";
				case ".tga": return "Image";
				case ".svg": return "Image";
				case ".ico": return "Image";
				case ".xls": return "Table";
				default:
					return "File";
			}
		}

		public void ExtractAttachments(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			var ext = Path.GetExtension(shortResourceNameWithExtension);

			switch (ext.ToLowerInvariant())
			{
				case ".eml":
				case ".msg": ExtractAttachmentsFromEmlOrMsg(input, shortResourceNameWithExtension, handler); break;
				case ".mbox": ExtractAttachmentsFromMbox(input, shortResourceNameWithExtension, handler); break;
				case ".ost":
				case ".pst": ExtractAttachmentsFromOstOrPst(input, shortResourceNameWithExtension, handler); break;
				default:
					throw new NotSupportedException($"Input type not supported {ext?.ToUpperInvariant()}");
			}
		}

		void ParseEmlOrMsg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var mail = MapiHelper.GetMapiMessageFromStream(input))
			{
				if (mail.Body != null)
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".txt")))
					using (var writer = new StreamWriter(output))
						writer.Write(mail.Body);

				if (mail.BodyRtf != null)
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".rtf")))
					using (var writer = new StreamWriter(output))
						writer.Write(mail.BodyRtf);

				if (mail.BodyHtml != null)
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".html")))
					using (var writer = new StreamWriter(output))
						writer.Write(mail.BodyHtml);

				if (mail.Attachments != null)
					foreach (var attachment in mail.Attachments)
					{
						using (var output = handler.CreateOutputStream(Path.GetFileName(attachment.LongFileName)))
							attachment.Save(output);
					}
			}
		}

		void ParseMbox(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			Convert(input, shortResourceNameWithExtension, handler, ".msg");
		}

		void ParseOstOrPst(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			Convert(input, shortResourceNameWithExtension, handler, ".msg");
		}

		private void ExtractAttachmentsFromEmlOrMsg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var mail = MapiHelper.GetMapiMessageFromStream(input))
			{
				if (mail.Attachments != null)
					foreach (var attachment in mail.Attachments)
					{
						using (var output = handler.CreateOutputStream(shortResourceNameWithExtension+"_"+Path.GetFileName(attachment.LongFileName)))
							attachment.Save(output);
					}
			}
		}

		private void ExtractAttachmentsFromMbox(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var reader = new MboxrdStorageReader(input, false))
			{
				for (int i = 0; i < reader.GetTotalItemsCount(); i++)
				{
					var message = reader.ReadNextMessage();

					foreach (var attachment in message.Attachments)
					{
						using (var output = handler.CreateOutputStream(shortResourceNameWithExtension + "_" + Path.GetFileName(attachment.Name)))
							attachment.Save(output);
					}
				}
			}
		}

		private void ExtractAttachmentsFromOstOrPst(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var personalStorage = PersonalStorage.FromStream(input))
			{
				HandleFolderAndSubfolders(mapiMessage =>
				{
					foreach (var attachment in mapiMessage.Attachments)
					{
						using (var output = handler.CreateOutputStream(shortResourceNameWithExtension + "_" + Path.GetFileName(attachment.LongFileName)))
							attachment.Save(output);
					}
				}, personalStorage.RootFolder);
			}
		}

		public void ExtractAttachmentsAsJpgImages(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			var ext = Path.GetExtension(shortResourceNameWithExtension);

			switch (ext.ToLowerInvariant())
			{
				case ".eml":
				case ".msg": ExtractAttachmentsAsJpgFromEmlOrMsg(input, shortResourceNameWithExtension, handler); break;
				case ".mbox": ExtractAttachmentsAsJpgFromMbox(input, shortResourceNameWithExtension, handler); break;
				case ".ost":
				case ".pst": ExtractAttachmentsAsJpgFromOstOrPst(input, shortResourceNameWithExtension, handler); break;
				default:
					throw new NotSupportedException($"Input type not supported {ext?.ToUpperInvariant()}");
			}
		}

		private void ExtractAttachmentsAsJpgFromEmlOrMsg(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var mail = MapiHelper.GetMailMessageFromStream(input))
			{
				if (mail.Attachments != null)
				{
					foreach (var attachment in mail.Attachments)
					{
						var type = attachment.ContentType.MediaType;

						START:
						switch (type)
						{
							case "message/rfc822":
							case "ost":
							case "pst":
							case "mbox":
							case "eml":
							case "msg":
								{
									var ms = new MemoryStream();
									attachment.Save(ms);
									ms.Position = 0;

									var name = Path.GetFileName(attachment.Name);

									if (attachment.ContentType.MediaType == "message/rfc822")
										name += ".msg";

									Convert(ms, name, handler, "jpg");
									break;
								}
							case "png":
							case "jpg":
							case "jpeg":
							case "tiff":
							case "bmp":
							case "svg":
								{
									using (var output = handler.CreateOutputStream(attachment.Name))
										attachment.Save(output);
									break;
								}
							case "application/vnd.ms-outlook":
							case "application/octet-stream":
								{
									var extension = Path.GetExtension(attachment.ContentType.Name)?.ToLowerInvariant();

									if (extension.IsNullOrWhiteSpace())
										break;
									else
									{
										type = extension.Substring(1);
										goto START;
									}
								}
							case "text/plain":
							case "application/msword":
							default:
								{
									try
									{
										var ms = new MemoryStream();
										attachment.Save(ms);
										ms.Position = 0;
										var document = new Document(ms);

										var saveOptions = new Aspose.Words.Saving.ImageSaveOptions(SaveFormat.Jpeg);

										if (document.PageCount == 1)
										{
											using (var output = handler.CreateOutputStream(shortResourceNameWithExtension + "_" + attachment.Name, ".jpg"))
												document.Save(output, saveOptions);
										}
										else
										{
											for (int i = 0; i < document.PageCount; i++)
											{
												saveOptions.PageSet = new Words.Saving.PageSet(i);

												using (var output = handler.CreateOutputStream(shortResourceNameWithExtension + "_" + attachment.Name, ".jpg", i))
													document.Save(output, saveOptions);
											}
										}

										break;
									}
									catch (NotSupportedException ex)
									{
										Logger.LogError(ex, "Unsupported email attachment", "Parser",  attachment.ContentType.MediaType);
										break;
									}
									catch (Aspose.Words.UnsupportedFileFormatException ex)
									{
										Logger.LogError(ex, "Unsupported email attachment", "Parser", attachment.ContentType.MediaType);
										break;
									}
								}
						}
					}
				}
			}
		}

		private void ExtractAttachmentsAsJpgFromMbox(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var reader = new MboxrdStorageReader(input, false))
			{
				for (int i = 0; i < reader.GetTotalItemsCount(); i++)
				{
					var message = reader.ReadNextMessage();

					foreach (var attachment in message.Attachments)
					{
						using (var output = handler.CreateOutputStream(shortResourceNameWithExtension + "_" + Path.GetFileName(attachment.Name)))
							attachment.Save(output);
					}
				}
			}
		}

		private void ExtractAttachmentsAsJpgFromOstOrPst(Stream input, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var personalStorage = PersonalStorage.FromStream(input))
			{
				HandleFolderAndSubfolders(mapiMessage =>
				{
					foreach (var attachment in mapiMessage.Attachments)
					{
						using (var output = handler.CreateOutputStream(shortResourceNameWithExtension + "_" + Path.GetFileName(attachment.LongFileName)))
							attachment.Save(output);
					}
				}, personalStorage.RootFolder);
			}
		}
	}
}
