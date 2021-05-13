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
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    public class AsposeEmailRedactionController : AsposeEmailBaseApiController
    {
        public AsposeEmailRedactionController(ILogger<AsposeEmailRedactionController> logger, 
			IConfiguration config, 
			IStorageService storageService,
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		/// <summary>
		/// Redact documents
		/// </summary>
		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("Redact")]
		public Task<Response> Redact(string outputType, string searchQuery, string replaceText, bool caseSensitive, bool text, bool comments, bool metadata)
		{
			return Process(RedactionApp, (service, handler, files) =>
			{
				if (outputType.IsNullOrEmpty())
					throw new BadRequestException("No output type provided");

				if (replaceText.IsNullOrEmpty())
					throw new BadRequestException("No replace text provided");

                foreach (var item in files)
                {
					using (var input = new MemoryStream(item.Value))
						service.Redact(input, item.Key, handler, searchQuery, replaceText, caseSensitive, text, comments, metadata, outputType);
				}
			});
		}
    }
}
