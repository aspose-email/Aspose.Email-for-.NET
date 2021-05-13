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
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    /// <summary>
    ///Implements the AsposeEmailWatermarkController Class to add or remove watermarks
    /// </summary>
    public class AsposeEmailWatermarkController : AsposeEmailBaseApiController
	{
        public AsposeEmailWatermarkController(ILogger<AsposeEmailWatermarkController> logger, 
			IConfiguration config,
			IStorageService storageService,
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

        public enum WatermarkActionType
		{
			text,
			image,
			remove
		}

		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("ProcessWatermark")]
		public Task<Response> ProcessWatermark(WatermarkActionType type, string textColor = null, string fontFamily = "Arial")
		{
			return Process(WatermarkApp, (service, handler, files) =>
			{
				var text = ReadAndRemoveAsText(files, "text");

				byte[] image = null;

				if (type == WatermarkActionType.image)
				{
					image = ReadAndRemoveAsBytes(files, files.Last().Key);

					if (image == null)
						throw new BadRequestException("Image not provided");

					try
					{
						var format = Aspose.Imaging.Image.GetFileFormat(new MemoryStream(image));

						if (format == Aspose.Imaging.FileFormat.Undefined)
							throw new BadRequestException("Image not provided");
					}
					catch(BadRequestException)
					{
						throw;
					}
					catch (Exception)
					{
						throw new BadRequestException("Image not provided");
					}
				}

				foreach(var pair in files)
				{
					var name = Path.GetFileName(pair.Key);

					using (var input = new MemoryStream(pair.Value))
					{
						switch (type)
						{
							case WatermarkActionType.text:
								if (text.IsNullOrWhiteSpace())
									throw new BadRequestException("Text not provided");
								if (textColor.IsNullOrWhiteSpace())
									throw new BadRequestException("Text Color not provided");
								if (fontFamily.IsNullOrWhiteSpace())
									throw new BadRequestException("Font family not provided");

								service.AddTextWatermark(input, name, text, fontFamily, textColor, handler);
								break;
							case WatermarkActionType.image:
								service.AddImageWatermark(input, name, image, handler);
								break;
							case WatermarkActionType.remove:
								service.RemoveWatermark(input, name, handler);
								break;
							default:
								throw new NotSupportedException("Unknown watermark actino type "+ type);
						}
					}
				}
			});
		}
	}
}
