namespace Aspose.Email.Live.Demos.UI.Models
{
    ///<Summary>
    /// EmailHeadersAnalyzerResponse class to get or set Email Headers Analyzer Response properties
    ///</Summary>
    public class EmailHeadersAnalyzerResponse
    {
		///<Summary>
		/// get or set StatusCode
		///</Summary>
		public int StatusCode { get; set; }

		///<Summary>
		/// get or set Status
		///</Summary>
		public string Status { get; set; }

		///<Summary>
		/// get or set RawHeaders
		///</Summary>
		public string RawHeaders { get; set; }

		///<Summary>
		/// get or set Traces
		///</Summary>
		public MailTraceInfo[] Traces { get; set; }
    }
}
