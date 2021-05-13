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
    ///<Summary>
    /// AsposeEmailParserController class to parse email file
    ///</Summary>
    public class AsposeEmailParserController : AsposeEmailBaseApiController
    {
        public AsposeEmailParserController(ILogger<AsposeEmailParserController> logger,
			IConfiguration config, 
			IStorageService storageService, 
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("Parse")]
		public Task<Response> Parse()
		{
			return Process(ParserApp, (service, handler, files) =>
			{
                foreach (var item in files)
                {
					using (var input = new MemoryStream(item.Value))
						service.Parse(input, item.Key, handler);
				}
			});
		}
    }
}
