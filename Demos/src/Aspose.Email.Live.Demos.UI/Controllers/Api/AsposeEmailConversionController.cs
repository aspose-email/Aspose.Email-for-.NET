using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Filters;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    ///<Summary>
    /// AsposeEmailConversionController class to convert email file to different format
    ///</Summary>
    public partial class AsposeEmailConversionController : AsposeEmailBaseApiController
	{
        public AsposeEmailConversionController(ILogger<AsposeEmailConversionController> logger,
			IConfiguration config, 
			IStorageService storageService, 
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		///<Summary>
		/// Convert method for uploaded file
		///</Summary>
		[HttpGet("ConvertAndDownload")]
		public async Task<IActionResult> ConvertAndDownload(string file, string folderName, string outputType)
		{
			// TODO: remove intermediate result
			var result = await Process(ConversionApp, folderName, file, (service, handler, files) =>
			{
				if (outputType.IsNullOrWhiteSpace())
					throw new BadRequestException("Output Type not provided");

				foreach (var pair in files)
				{
					using (var input = new MemoryStream(pair.Value))
						service.Convert(input, Path.GetFileName(pair.Key), handler, outputType);
				}
			});

			if (result.StatusCode == 200)
			{
				using (var input = await StorageService.ReadFile(result.FolderName, result.FileName))
				{
					var httpResponce = new HttpResponseMessage(HttpStatusCode.OK)
					{
						Content = new ByteArrayContent(input.ToArray())
					};

					httpResponce.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
					{
						FileName = result.FileName
					};

					httpResponce.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

					return httpResponce.ToActionResult();
				}
			}

			return BadRequest();
		}

		///<Summary>
		/// Convert method
		///</Summary>
		//[ValidateAntiForgeryToken]
		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("Convert")]
		public Task<Response> Convert(string outputType)
		{
			return Process(ConversionApp, (service, handler, files) =>
			{
				if (outputType.IsNullOrWhiteSpace())
					throw new BadRequestException("Output Type not provided");

				foreach (var pair in files)
				{
					using (var input = new MemoryStream(pair.Value))
						service.Convert(input, Path.GetFileName(pair.Key), handler, outputType);
				}
			});
		}
	}
}
