namespace Aspose.Email.Live.Demos.UI.Models
{
	public class MailTraceInfo
    {
		public string Domain { get; set; }
		public string Ip { get; set; }
		public Aspose.Email.HeadersEngine.IpLocation Location { get; set; }
		public Aspose.Email.HeadersEngine.WhoIsResponse WhoIs { get; set; }

		public MailTraceInfo(Aspose.Email.HeadersEngine.MailTraceInfo trace)
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
