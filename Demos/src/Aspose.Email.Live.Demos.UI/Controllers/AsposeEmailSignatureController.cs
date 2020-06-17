using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Mapi;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using File = System.IO.File;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	///<Summary>
	/// AsposeEmailSignatureController class to sign email file
	///</Summary>
	public class AsposeEmailSignatureController : EmailBase
	{
		private const int ImageHeight = 100;
		///<Summary>
		/// initialize AsposeEmailSignatureController class 
		///</Summary>
		public AsposeEmailSignatureController()
		{
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
		}

		///<Summary>
		/// Sign documents
		///</Summary>
		[MimeMultipart]
		[HttpPost]
		[AcceptVerbs("GET", "POST")]
		public async Task<Response> Sign(string outputType, string signatureType)
		{
			(string folder, string[] files) = (null, null);

			try
			{
				(folder, files) = await UploadFiles();

				var image = ReadAndRemoveAsText(ref files, folder, "image");
				var text = ReadAndRemoveAsText(ref files, folder, "text");
				var textColor = ReadAndRemoveAsText(ref files, folder, "textColor");

				if (image == null && text == null)
					image = Convert.ToBase64String(ReadAndRemoveAsBytes(ref files, folder, files.Last()));

				var processor = new CustomSingleOrZipFileProcessor()
				{
					
					CustomFilesProcessMethod = (string[] inputFilePaths, string outputFolderPath) =>
					{
						for (int i = 0; i < inputFilePaths.Length; i++)
						{
							var inputFilePath = inputFilePaths[i];


							var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

							if (mail == null)
								throw new Exception("Invalid file format");

							switch (signatureType)
							{
								case "image":
								case "drawing":
									mail = AddDrawing(mail, image);
									break;
								case "text":
									mail = AddText(mail, text, textColor);
									break;
								default:
									throw new Exception("Unknown signature type.");
							}

							if (outputType.Equals("MSG", StringComparison.OrdinalIgnoreCase))
								mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Signed.msg"));
							else
								mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Signed.eml"), new EmlSaveOptions(MailMessageSaveType.EmlFormat));
						}
					}
				};

				return processor.Process(folder, files);
			}
			catch (Exception ex)
			{
				
				return new Response() { Status = "200", Text = "Error on processing file" };
			}
		}

		///<Summary>
		/// Sign method to sign email file
		///</Summary>
		public Task<Response> Sign(string inputFileName, string inputFolderName, string outputType, string signatureType, dynamic signatureObject)
		{
			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomFileProcessMethod= (string inputFilePath, string outputFolderPath) =>
				{
					var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

					if (mail == null)
						throw new Exception("Invalid file format");

					switch (signatureType)
					{
						case "drawing":
							mail = AddDrawing(mail, (string)signatureObject["DrawingImage"]);
							break;
						case "text":
							mail = AddText(mail, (string)signatureObject["Text"], (string)signatureObject["TextColor"]);
							break;
						case "image":
							mail = AddImage(mail, (string)signatureObject["ImageFolder"], (string)signatureObject["ImageFile"]);
							break;
						default:
							throw new Exception("Unknown signature type.");
					}

					if (outputType.Equals("MSG", StringComparison.OrdinalIgnoreCase))
						mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Signed.msg"));
					else
						mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Signed.eml"), new EmlSaveOptions(MailMessageSaveType.EmlFormat));
				}
			};



			return Task.FromResult(processor.Process(inputFolderName, inputFileName));
		}

		private MapiMessage AddDrawing(MapiMessage mail, string imageBase64)
		{
			var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");

			var element = htmlDocument.CreateElement("Signature");
			element.InnerHTML = "<img Src=\"data: image / png; base64, " + imageBase64 + "\" />";
			htmlDocument.Body.AppendChild(element);

			var folderPath = Path.Combine(Config.Configuration.OutputDirectory, Guid.NewGuid().ToString());
			var filePath = Path.Combine(folderPath, "Merged.html");

			htmlDocument.Save(filePath);

			var content = System.IO.File.ReadAllText(filePath);

			System.IO.File.Delete(filePath);
			Directory.Delete(folderPath);

			mail.SetBodyContent(content, BodyContentType.Html);

			return mail;
		}

		private MapiMessage AddText(MapiMessage mail, string text, string textColor)
		{
			var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");

			var element = htmlDocument.CreateElement("Signature");
			element.InnerHTML = "<p style = \"color:"+ textColor + "\";>" + text + "</p>";

			htmlDocument.Body.AppendChild(element);

			var folderPath = Path.Combine(Config.Configuration.OutputDirectory, Guid.NewGuid().ToString());
			var filePath = Path.Combine(folderPath, "Merged.html");

			htmlDocument.Save(filePath);

			var content = System.IO.File.ReadAllText(filePath);

			System.IO.File.Delete(filePath);
			Directory.Delete(folderPath);

			mail.SetBodyContent(content, BodyContentType.Html);

			return mail;
		}

		

		private MapiMessage AddImage(MapiMessage mail, string imageFoler, string imageFile)
		{
			var imageBase64 = Convert.ToBase64String(System.IO.File.ReadAllBytes(Path.Combine(Config.Configuration.WorkingDirectory, imageFoler, imageFile)));
			var htmlDocument = new Aspose.Html.HTMLDocument(mail.BodyHtml, "");

			var element = htmlDocument.CreateElement("Signature");
			element.InnerHTML = "<img Src=\"data: image / png; base64, " + imageBase64 + "\" />";
			htmlDocument.Body.AppendChild(element);

			var folderPath = Path.Combine(Config.Configuration.OutputDirectory, Guid.NewGuid().ToString());
			var filePath = Path.Combine(folderPath, "Merged.html");

			htmlDocument.Save(filePath);

			var content = System.IO.File.ReadAllText(filePath);

			System.IO.File.Delete(filePath);
			Directory.Delete(folderPath);

			mail.SetBodyContent(content, BodyContentType.Html);

			return mail;
		}
	}
}
