using System.IO;

namespace Aspose.Email.Live.Demos.UI.Services
{
	/// <summary>
	/// Interface for saving output data and files during Service processing or for results
	/// </summary>
	public interface IOutputHandler
    {
		/// <summary>
		/// Create new output stream for resource by name.
		/// Should recreate if resource exists.
		/// </summary>
		/// <param name="name">Name of resource</param>
		/// <returns>Empty stream for writing</returns>
		Stream CreateOutputStream(string name);

		/// <summary>
		/// Open resource if exists or throw
		/// </summary>
		/// <param name="name">Name of resource</param>
		/// <returns>New stream for read</returns>
		Stream OpenReadStream(string name);

		/// <summary>
		/// Remove resource by name if exists
		/// </summary>
		/// <param name="name">Name of resource</param>
		void RemoveResource(string name);

		/// <summary>
		/// Builds name for partial output file. Preferred for using with <see cref="CreateOutputStream(string)"/>
		/// </summary>
		/// <param name="nameWithoutExtension">File name</param>
		/// <param name="extension">File extension</param>
		/// <param name="partIndex">Part index</param>
		/// <returns>Built name</returns>
		string BuildName(string nameWithoutExtension, string extension, int? partIndex = null);
    }
}
