using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Email.Live.Demos.UI.Models
{
	public class EmailHeadersAnalyzerResponse
	{
		///<Summary>
		/// get or set StatusCode
		///</Summary>
		public int StatusCode { get; set; }
		///<Summary>
		/// get or set Status
		///</Summary>
		public String Status { get; set; }
		///<Summary>
		/// get or set RawHeaders
		///</Summary>
		public string RawHeaders { get; set; }
		///<Summary>
		/// get or set Traces
		///</Summary>
		public MailTraceInfo[] Traces { get; set; }
	}
	public class MailTraceInfo
	{
		public string Domain { get; set; }
		public string Ip { get; set; }
		public Email.HeadersEngine.IpLocation Location { get; set; }
		public Email.HeadersEngine.WhoIsResponse WhoIs { get; set; }

		public MailTraceInfo(Email.HeadersEngine.MailTraceInfo trace)
		{
			Domain = trace.Domain;
			Location = trace.Location;
			WhoIs = trace.WhoIs;
			Ip = trace.Ip?.ToString();
		}

		public MailTraceInfo()
		{
		}
	}
}