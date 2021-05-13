using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    using Aspose.Email.Live.Demos.UI.Services;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using System.Threading.Tasks;

    ///<Summary>
    /// AsposeViewerController class to get document page
    ///</Summary>
    public class AsposeEmailViewerController : AsposeEmailBaseApiController
	{
        public AsposeEmailViewerController(ILogger<AsposeEmailViewerController> logger,
			IConfiguration config, 
			IStorageService storageService,
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		///<Summary>
		/// DocumentPages method to get document pages
		///</Summary>
		[HttpGet("DocumentPages")]
		public async Task<List<string>> BuildAndStoreDocumentPages(string file, string folderName)
		{
			var logMessage = "ControllerName = AsposeViewerController, MethodName = DocumentPages, Folder = " + folderName;
			var output = new List<string>();

			try
			{
				var viewerImages = await GetDocumentPages(file, folderName);

				var names = viewerImages.GetOrderedPageNames();
				var format = "Page {0} ({1})";

				output.Add(names.Length.ToString());
				var imagesFolderName = Guid.NewGuid().ToString();

				for (int k = 0; k < names.Length; k++)
				{
					using (var fileStream = viewerImages.OpenReadStream(names[k]))
					{
						var imageName = string.Format(format, k + 1, names[k]);

						await StorageService.SaveFile(fileStream, imagesFolderName, imageName);

						fileStream.Position = 0;

						using (var image = Aspose.Imaging.Image.Load(fileStream))
						{
							var token = new JObject()
							{
								{ "ImageName", imageName },
								{ "ImageSize", $"{image.Width} x {image.Height}" },
								{ "ImageFolderName", imagesFolderName }
							};

							output.Add(token.ToString());
						}
					}
				}

				Logger.LogInformation(logMessage, AsposeEmail, file);
			}
			catch(BadRequestException ex)
			{
				throw;
			}
			catch (Exception ex)
			{
				Logger.LogError(ex, logMessage, AsposeEmail, file);
				throw;
			}

			return output;
		}

		private async Task<ViewerImagesOutputHandler> GetDocumentPages(
			string file,
			string folderName)
		{
			if (file.IsNullOrWhiteSpace())
				throw new BadRequestException("File name not provided");

			if (folderName.IsNullOrWhiteSpace())
				throw new BadRequestException("Folder name not provided");

			var filename = file;

			// Page file name format
			var format = "Page {0} ({1})";

			// We will save all outputs in viewer format
			var viewerImagesHandler = new ViewerImagesOutputHandler(format + ".jpg");

			// Convert main file to jpg images
			using (var input = await StorageService.ReadFile(folderName, filename))
			{
				EmailService.Convert(input, filename, viewerImagesHandler, "jpg");

				// TODO: add full support for downloading attachments images (not only on viewing)
				//input.Position = 0;
				//service.ExtractAttachmentsAsJpgImages(input, filename, viewerImagesHandler);
			}
			

			//await viewerImagesHandler.SaveAllImagesInStorage();

			return viewerImagesHandler;
		}
		///<Summary>
		/// DownloadDocument method to download document
		///</Summary>
		[HttpGet("DownloadDocument")]
		public async Task<IActionResult> DownloadDocument(string file, string folderName)
		{
			string logMsg = "ControllerName = AsposeEmailViewerController, MethodName = DownloadDocument, Folder = " + folderName;

			try
			{
				using (var stream = await StorageService.ReadFile(folderName, file))
				{
					Logger.LogInformation(logMsg, AsposeEmail, file);

					using (var ms = new MemoryStream())
					{
						await stream.CopyToAsync(ms);

						var result = new HttpResponseMessage(HttpStatusCode.OK)
						{
							Content = new ByteArrayContent(ms.ToArray())
						};

						result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
						{
							FileName = file
						};

						result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

						return result.ToActionResult();
					}
				}
			}
			catch (Exception x)
			{
				Logger.LogError(x, logMsg, AsposeEmail, file);
				throw;
			}
		}

		///<Summary>
		/// PageImage method to get page image
		///</Summary>
		[HttpGet("PageImage")]
		public async Task<IActionResult> PageImage(string imageFolderName, string imageFileName)
		{
			return (await GetImageFromStorage(imageFolderName, imageFileName)).ToActionResult();
		}

		private async Task<HttpResponseMessage> GetImageFromStorage(string imageFolderName, string imageFileName)
        {
            if (!await StorageService.IsFileExists(imageFolderName, imageFileName))
            {
				return new HttpResponseMessage(HttpStatusCode.NoContent);
            }

			HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            var fileExt = Path.GetExtension(imageFileName).ToLower();
			string contentType = null;
            switch (fileExt)
            {
				case ".png":
				case ".apng":
                    contentType = "image/png";
					break;
                case ".jpg":
                case ".jpeg":
                    contentType = "image/jpeg";
                    break;
                case ".gif":
                    contentType = "image/gif";
                    break;
                case ".webp":
                    contentType = "image/webp";
                    break;
                case ".svg":
                    contentType = "image/svg+xml";
                    break;
			}

            if (contentType != null)
            {
				using (var imageStream = await StorageService.ReadFile(imageFolderName, imageFileName))
				{
					using (var ms = new MemoryStream())
					{
						await imageStream.CopyToAsync(ms);
						result.Content = new ByteArrayContent(ms.ToArray());
						result.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
						return result;
					}
				}
            }

			using (var imageStream = await StorageService.ReadFile(imageFolderName, imageFileName))
			{
				var image = System.Drawing.Image.FromStream(imageStream);
				var memoryStream = new MemoryStream();

				image.Save(memoryStream, ImageFormat.Jpeg);

				result.Content = new ByteArrayContent(memoryStream.ToArray());
				result.Content.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg");
			}

			return result;
		}
	}
}
