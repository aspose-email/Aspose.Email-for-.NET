using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Filters;
using Aspose.Email.Live.Demos.UI.Helpers;
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
    public class AsposeEmailAssemblyController : AsposeEmailBaseApiController
	{
        public AsposeEmailAssemblyController(ILogger<AsposeEmailAssemblyController> logger,
			IConfiguration config,
			IStorageService storageService, 
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

		[DisableFormValueModelBinding]
		[RequestSizeLimit(AppConstants.MaxFileSize)]
		[RequestFormLimits(MultipartBodyLengthLimit = AppConstants.MaxFileSize)]
		[HttpPost("Assemble")]
		public Task<Response> Assemble(string sourceName, int tableIndex = 0, string delimiter = ",")
		{
			return Process(AssemblyApp, (service, handler, files) =>
			{
				if (files.Count < 2)
					throw new BadRequestException("Request doesn't contain template or dataSource");

				if (string.IsNullOrWhiteSpace(sourceName))
					throw new BadRequestException("Invalid source name");

				var templatePair = files.ElementAt(0);
				var dataSourcePair = files.ElementAt(1);

				var dataTable = AssemblyDataHelper.PrepareDataTable(dataSourcePair.Key, new MemoryStream(dataSourcePair.Value), sourceName, tableIndex, delimiter);

				if (dataTable == null)
					throw new BadRequestException("Can't process the data source");

				using (var input = new MemoryStream(templatePair.Value))
					service.AssembleFile(input, Path.GetFileName(templatePair.Key), handler, dataTable);
			});
		}
	}
}
