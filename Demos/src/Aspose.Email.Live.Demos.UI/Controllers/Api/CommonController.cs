using Aspose.Email.Live.Demos.UI.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    using Aspose.Email.Live.Demos.UI.Config;
    using Aspose.Email.Live.Demos.UI.Filters;
    using Aspose.Email.Live.Demos.UI.Helpers;
    using Aspose.Email.Live.Demos.UI.Services;
    using Aspose.Email.Live.Demos.UI.Services.Email;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using System.IO.Compression;

    /// <summary>
    /// Common API controller.
    /// </summary>
    public class CommonController : AsposeEmailBaseApiController
	{
		public CommonController(ILogger<CommonController> logger,
			IConfiguration config,
			IStorageService storageService,
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
		}

		/// <summary>
		/// Sends back specified file from specified folder inside OutputDirectory.
		/// </summary>
		/// <param name="folder">Folder inside OutputDirectory.</param>
		/// <param name="file">File.</param>
		/// <returns>HTTP response with file.</returns>
		[ApiExplorerSettings(IgnoreApi = true)]
		[NonAction]
		private async Task<HttpResponseMessage> Download(string folder, string file)
		{
			//var pathProcessor = new PathProcessor(folder, file: file);
			//var stream = new FileStream(pathProcessor.DefaultOutFile, FileMode.Open, FileAccess.Read, FileShare.None, bufferSize: 4096, useAsync: true);
			var ms = new MemoryStream();

			using (var stream = await StorageService.ReadFile(folder, file))
			{
				await stream.CopyToAsync(ms);
				ms.Position = 0;
			}

			var result = new HttpResponseMessage(HttpStatusCode.OK)
			{
				Content = new StreamContent(ms)
			};

			result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
			result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
			{
				FileName = file
			};
			return result;
		}

        /// <summary>
        /// Uploads file.
        /// Returns upload id and filename.
        /// </summary>
        /// <returns>File upload result.</returns>
        //[MimeMultipart]
        [DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("UploadFile")]
		public async Task<FileSafeResult> UploadFile()
		{
			return (await UploadMultipleFiles()).First();
		}

		/// <summary>
		/// Uploads files.
		/// Returns collection of upload id and filename.
		/// </summary>
		/// <returns>Files upload results.</returns>
		//[MimeMultipart]
		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("UploadMultipleFiles")]
		public async Task<IEnumerable<FileSafeResult>> UploadMultipleFiles()
		{
			var files = await UploadFiles();

			var folder = Guid.NewGuid().ToString();

			return await Task.WhenAll(files.Select(async item =>
			{
				await StorageService.SaveFile(item.Value, folder, item.Key);

				return new FileSafeResult()
				{
					id = folder,
					FileName = item.Key,
					UploadFileName = item.Key,
					idUpload = folder
				};
			}).ToArray());
		}

		/// <summary>
		/// Downloads file with specified upload id and filename.
		/// </summary>
		/// <param name="id">Download id.</param>
		/// <param name="file">File name.</param>
		/// <returns>Binary stream mime type application/octet-stream</returns>	
		[HttpGet("DownloadFile/{id}")]
		public async Task<IActionResult> DownloadFile(string id, string file) => new HttpResponseMessageResult(await Download(id, file));
		
		/// <summary>
		/// Downloads the files.
		/// </summary>
		/// <param name="id">The identifier.</param>
		/// <param name="files">The file names.</param>
		/// <returns> Binary stream mime type application/octet-stream </returns>	
		[HttpGet("DownloadFiles")]
        public async Task<IActionResult> DownloadFiles(string id, [FromUri] string[] files)
        {
			using (var stream = new MemoryStream())
			{
				using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
				{
					for (int i = 0; i < files.Length; i++)
					{
						var name = files[i];
						var file = archive.CreateEntry(name);
						using (var entryStream = file.Open())
						{
							using (var fileStream = await StorageService.ReadFile(id, name))
								await fileStream.CopyToAsync(entryStream);
						}
					}
				}

				stream.Position = 0;

				var result = new HttpResponseMessage(HttpStatusCode.OK)
				{
					Content = new StreamContent(stream)
				};

				result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
				result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
				{
					FileName = Guid.NewGuid().ToString()
				};

				return result.ToActionResult();
			}
        }
	}
}
