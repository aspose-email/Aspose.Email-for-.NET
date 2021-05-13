namespace Aspose.Email.Live.Demos.UI.Services
{
    public interface IAppResources
    {
        string GetResourceOrDefault(string resourceName, string defaultValue = null);
        bool ContainsResource(string resourceName);
    }
}
