using Aspose.Email.HeadersEngine;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
    public partial class EmailService
	{
		public async Task<Models.MailTraceInfo[]> ExtractHeaders(Stream messageStream, bool skipLocal)
		{
			var analyzer = new MailHeaderAnalyzer();
			var trace = await analyzer.GetMailTraceInfo(messageStream, !skipLocal);
			return trace.Select(x => new Models.MailTraceInfo(x)).ToArray();
		}
	}
}
