using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Filters;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    ///<Summary>
    /// AsposeEmailHeadersController class to analyze email file and headers
    ///</Summary>
    public class AsposeEmailHeadersController : AsposeEmailBaseApiController
    {
        public AsposeEmailHeadersController(ILogger<AsposeEmailHeadersController> logger, 
			IConfiguration config, 
			IStorageService storageService,
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		///<Summary>
		/// AnalyzeEmailFile method to analyze email file
		///</Summary>
		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("AnalyzeEmailFiles")]
		public Task<EmailHeadersAnalyzerResponse> AnalyzeEmailFiles(bool skipLocal)
		{
			return UploadAndHandle("Headers", async (IDictionary<string, byte[]> files) =>
			{
				if (files.Count < 1)
					throw new BadRequestException("No files provided");

				var file = files.First();

				using (var stream = new MemoryStream(file.Value))
				{
					var responce = new EmailHeadersAnalyzerResponse()
					{
						Status = "OK",
						StatusCode = 200,
					};

					responce.Traces = await EmailService.ExtractHeaders(stream, skipLocal);

					return responce;
				}
			},
			(ex, files) => new EmailHeadersAnalyzerResponse() { StatusCode = 400, Status = ex.Message },
			(ex, files) => new EmailHeadersAnalyzerResponse() { StatusCode = 500, Status = "Error on processing: " + ex.Message });
		}
    }
}
