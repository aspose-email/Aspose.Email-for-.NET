using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Filters;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Aspose.Words;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    ///<Summary>
    /// AsposeEmailSearchController class to perfrom search in email file
    ///</Summary>
    public class AsposeEmailSearchController : AsposeEmailBaseApiController
	{
		// Yellow Color
		private static readonly Color HiglightedColor = Color.FromArgb(246, 247, 146);

        public AsposeEmailSearchController(ILogger<AsposeEmailSearchController> logger, 
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
		[HttpPost("Search")]
		public async Task<Response> Search(string query)
		{
			if (string.IsNullOrEmpty(query))
			{
				return new Response
				{
					Status = "No query provided",
					StatusCode = 400,
					FileProcessingErrorCode = FileProcessingErrorCode.WrongRegExp
				};
			}

			if (!query.IsValidRegex())
			{
				return new Response
				{
					Status = "OK",
					StatusCode = 200,
					FileProcessingErrorCode = FileProcessingErrorCode.WrongRegExp
				};
			}

			return await Process(SearchApp, (service, handler, files) =>
			{
				var doc = new Document();
				var builder = new DocumentBuilder(doc);

				builder.MoveToSection(0);

				int allMatchesCounter = 0;

				foreach(var pair in files)
				{
					var sectionStart = builder.CurrentSection.Body.FirstParagraph;

					int fileMatchesCounter = 0;

					using (var input = new MemoryStream(pair.Value))
					{
						foreach (var match in service.Search(input, query))
						{
							var start = new Run(doc, match.TextBefore);
							var value = new Run(doc, match.TextMatch);
							var end = new Run(doc, match.TextAfter);

							value.Font.HighlightColor = HiglightedColor;

							builder.Writeln((match.MatchIndex + 1) + ":");

							builder.InsertNode(start);
							builder.InsertNode(value);
							builder.InsertNode(end);

							builder.Writeln();
							builder.Writeln();

							allMatchesCounter++;
							fileMatchesCounter++;
						}

						var endSection = builder.CurrentSection.Body.FirstParagraph;
						builder.MoveTo(sectionStart);

						builder.Writeln("File: " + Path.GetFileName(pair.Key));
						builder.Writeln("Matches found: " + fileMatchesCounter);
						builder.Writeln();
						builder.MoveTo(endSection);
					}
				}

				builder.MoveToSection(0);
				builder.Writeln("All Matches found: " + allMatchesCounter);
				builder.Writeln();

				using (var output = handler.CreateOutputStream("SearchResults.docx"))
					doc.Save(output, SaveFormat.Docx);
			});
		}
	}
}
