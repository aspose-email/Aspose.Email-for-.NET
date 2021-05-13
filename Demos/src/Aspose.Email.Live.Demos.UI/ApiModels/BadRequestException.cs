using System;
using System.Runtime.Serialization;

namespace Aspose.Email.Live.Demos.UI.Models
{
	public class BadRequestException : Exception
	{
		public BadRequestException()
		{
		}

		public BadRequestException(string message)
			: base(message)
		{
		}

		public BadRequestException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		protected BadRequestException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
