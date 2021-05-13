using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Filters;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    ///<Summary>
    /// AsposeEmailMergerController class to merge email file
    ///</Summary>
    public class AsposeEmailMergerController : AsposeEmailBaseApiController
	{
        public AsposeEmailMergerController(ILogger<AsposeEmailMergerController> logger,
			IConfiguration config, 
			IStorageService storageService, 
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("Merge")]
		public Task<Response> Merge()
		{
			return Process(MergerApp, (service, handler, files) =>
			{
				if (files.Count < 2)
					throw new BadRequestException("Required minimum two files for merging");

				if (files.Count > 10)
					throw new BadRequestException("10 files is maximum for merging. Please, remove excess files");

				var inputs = files.Select(x => { return ((Stream)new MemoryStream(x.Value), Path.GetFileName(x.Key)); });

				service.Merge(inputs, handler);
			});
		}
	}
}


