using System;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.HeadersEngine;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	///<Summary>
	/// AsposeEmailHeadersController class to analyze email file and headers
	///</Summary>
	public class AsposeEmailHeadersController : EmailBase
	{
		///<Summary>
		/// initialize AsposeEmailHeadersController class
		///</Summary>
		static AsposeEmailHeadersController()
        {
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
        }

        delegate Task<System.IO.Stream> DataProvider();
		///<Summary>
		/// AnalyzeRawHeaders method to analyze raw headers
		///</Summary>
		[HttpPost]        
        public async Task<EmailHeadersAnalyzerResponse> AnalyzeRawHeaders(bool skipLocal)
        {
            return await AnalyzeStream(skipLocal, async ()=> await Request.Content.ReadAsStreamAsync() );
        }

		///<Summary>
		/// AnalyzeEmailFile method to analyze email file
		///</Summary>
		[HttpPost]		
		public async Task<EmailHeadersAnalyzerResponse> AnalyzeEmailFiles(bool skipLocal)
		{
			var files = await UploadFiles();

			return await AnalyzeStream(skipLocal, async () => {
				string fullPath = System.IO.Path.Combine(Config.Configuration.WorkingDirectory, files.Item1, files.Item2[0]);

				return await Task.FromResult(System.IO.File.OpenRead(fullPath));
			});
		}

		///<Summary>
		/// AnalyzeEmailFile method to analyze email file
		///</Summary>
		[HttpGet]        
        public async Task<EmailHeadersAnalyzerResponse> AnalyzeEmailFile(string fileName, string folderName, bool skipLocal)
        {
            return await AnalyzeStream(skipLocal, async () => {
                string fullPath = System.IO.Path.Combine(Config.Configuration.WorkingDirectory, folderName, fileName);

                return await Task.FromResult(System.IO.File.OpenRead(fullPath));
            });
        }

        private static async Task<EmailHeadersAnalyzerResponse> AnalyzeStream(bool skipLocal, DataProvider dataProvider)
        {

            string statusValue = "OK";
            int statusCodeValue = 200;
			Models.MailTraceInfo[] items = null;

            try
            {
                using (var stream = await dataProvider())
                {
                    var analyzer = new MailHeaderAnalyzer();
                    var trace = await analyzer.GetMailTraceInfo(stream, !skipLocal);
                    items = trace.Select(x => new Models.MailTraceInfo(x)).ToArray();
                }
            }
            catch (Exception ex)
            {
                statusCodeValue = 500;
                statusValue = ex.Message;
            }

            return await Task.FromResult(new EmailHeadersAnalyzerResponse
            {
                Status = statusValue,
                StatusCode = statusCodeValue,
                Traces = items,
            });
        }
    }
}
