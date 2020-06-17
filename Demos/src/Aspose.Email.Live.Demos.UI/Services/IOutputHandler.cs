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
		/// Remove resource by name if exists
		/// </summary>
		/// <param name="name">Name of resource</param>
		void RemoveResource(string name);
    }
}
