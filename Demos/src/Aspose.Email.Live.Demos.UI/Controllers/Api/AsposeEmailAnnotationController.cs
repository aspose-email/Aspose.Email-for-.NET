using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Filters;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    public class AsposeEmailAnnotationController : AsposeEmailBaseApiController
    {
        public AsposeEmailAnnotationController(ILogger<AsposeEmailAnnotationController> logger,
			IConfiguration config, 
			IStorageService storageService,
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		///<Summary>
		/// Remove annotations method
		///</Summary>
		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("Remove")]
		public Task<Response> Remove()
		{
			return Process(AnnotationApp, (service, handler, files) =>
			{
                foreach (var pair in files)
                {
					using (var input = new MemoryStream(pair.Value))
						service.RemoveAnnotation(input, Path.GetFileName(pair.Key), handler);
				}
			});
		}
    }
}
