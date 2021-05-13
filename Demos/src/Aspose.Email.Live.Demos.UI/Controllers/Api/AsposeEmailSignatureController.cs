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
    ///<Summary>
    /// AsposeEmailSignatureController class to sign email file
    ///</Summary>
    public class AsposeEmailSignatureController : AsposeEmailBaseApiController
	{
        public AsposeEmailSignatureController(ILogger<AsposeEmailSignatureController> logger, 
			IConfiguration config, 
			IStorageService storageService, 
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

        public enum SignatureType
		{
			image,
			drawing,
			text
		}

		///<Summary>
		/// Sign documents
		///</Summary>
		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("Sign")]
		public Task<Response> Sign(string outputType, SignatureType signatureType)
		{
			return Process(SignatureApp, (service, handler, inputFilePaths) =>
			{
				if (outputType.IsNullOrWhiteSpace())
					throw new BadRequestException("Output type not provided");

				var image = ReadAndRemoveAsText(inputFilePaths, "image");
				var text = ReadAndRemoveAsText(inputFilePaths, "text");
				var textColor = ReadAndRemoveAsText(inputFilePaths, "textColor");

                foreach (var item in inputFilePaths)
                {
					var fileName = item.Key;

					using (var input = new MemoryStream(item.Value))
					{
						switch (signatureType)
						{
							case SignatureType.image:
							case SignatureType.drawing:
								service.AddDrawingSignature(input, fileName, image, outputType, handler);
								break;
							case SignatureType.text:
								service.AddTextSignature(input, fileName, text, textColor, outputType, handler);
								break;
							default:
								throw new Exception("Unknown signature type.");
						}
					}
				}
			});
		}
	}
}
