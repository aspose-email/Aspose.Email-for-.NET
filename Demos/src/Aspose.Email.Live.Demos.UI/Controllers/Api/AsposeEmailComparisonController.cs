using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Filters;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    public class AsposeEmailComparisonController : AsposeEmailBaseApiController
    {
        public AsposeEmailComparisonController(ILogger<AsposeEmailComparisonController> logger,
			IConfiguration config,
			IStorageService storageService, 
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		///<Summary>
		/// Merge method to merge word document
		///</Summary>
		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("Compare")]
		public Task<Response> Compare()
		{
			return Process(ComparisonApp, (service, handler, files) =>
			{
				if (files.Count < 2)
					throw new BadRequestException("Minimum 2 files should be provided");

				var file1Pair = files.ElementAt(0);
				var file2Pair = files.ElementAt(1);

				service.Compare(new MemoryStream(file1Pair.Value), new MemoryStream(file2Pair.Value), Path.GetFileName(file1Pair.Key), handler);
			});
		}
    }
}
