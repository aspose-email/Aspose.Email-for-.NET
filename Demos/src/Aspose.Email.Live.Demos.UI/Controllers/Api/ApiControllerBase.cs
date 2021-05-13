using Aspose.Email.Live.Demos.UI.Helpers;
using Microsoft.AspNetCore.Mvc;
using System.Web.Http;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	///<Summary>
	/// ApiControllerBase class to have base methods
	///</Summary>
	[ApiController]
	[Route("api/[controller]")]
	[DefaultExceptionFilter]
	public abstract class ApiControllerBase : ApiController
	{
		///<Summary>
		/// Aspose Email
		///</Summary>
		protected string AsposeEmail => "Aspose.Email";
		///<Summary>
		/// Aspose Conversion APP
		///</Summary>
		protected string ConversionApp => " Conversion";
		///<Summary>
		/// Aspose Watermark APP
		///</Summary>
		protected string WatermarkApp => " Watermark";
		///<Summary>
		/// Aspose Merger APP
		///</Summary>
		protected string MergerApp => " Merger";
		///<Summary>
		/// Aspose ParserApp APP
		///</Summary>
		protected string ParserApp => " Parser";
		///<Summary>
		/// Aspose Annotation APP
		///</Summary>
		protected string AnnotationApp => " Annotation";
		///<Summary>
		/// Aspose Redaction APP
		///</Summary>
		protected string RedactionApp => " Redaction";
		///<Summary>
		/// Aspose Editor APP
		///</Summary>
		protected string EditorApp => " Editor";
		///<Summary>
		/// Aspose Comparison APP
		///</Summary>
		protected string ComparisonApp => " Comparison";
		///<Summary>
		/// Aspose Assembly APP
		///</Summary>
		protected string AssemblyApp => " Assembly";
		///<Summary>
		/// Search
		///</Summary>
		protected string SearchApp => " Search";
		///<Summary>
		/// Signature App Property
		///</Summary>
		protected string SignatureApp => " Signature";
	}
}
