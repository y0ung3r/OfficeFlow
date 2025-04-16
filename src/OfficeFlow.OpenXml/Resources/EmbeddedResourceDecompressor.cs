using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using OfficeFlow.OpenXml.Resources.Exceptions;
using OfficeFlow.OpenXml.Resources.Interfaces;

namespace OfficeFlow.OpenXml.Resources;

public sealed class EmbeddedResourceDecompressor(Assembly assembly) : IResourceDecompressor
{
    public EmbeddedResourceDecompressor()
        : this(Assembly.GetCallingAssembly())
    { }

    private Stream GetResourceStream(string resourceName)
    {
        var resourcePath = assembly
            .GetManifestResourceNames()
            .FirstOrDefault(name => name.EndsWith(resourceName))
                ?? throw new ResourceNotFoundException(resourceName);

        return assembly.GetManifestResourceStream(resourcePath)
            ?? throw new ResourceNotLoadedException(resourceName);
    }

    public XDocument Decompress(string resourceName)
    {
        using var stream = 
            GetResourceStream(resourceName);
        
        using var zip = 
            new GZipStream(stream, CompressionMode.Decompress);
        
        using var reader = 
            new StreamReader(zip);

        return XDocument.Load(reader);
    }
}