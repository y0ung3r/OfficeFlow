using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using OfficeFlow.Word.OpenXml.Resources.Exceptions;
using OfficeFlow.Word.OpenXml.Resources.Interfaces;

namespace OfficeFlow.Word.OpenXml.Resources.Embedded
{
    public sealed class EmbeddedResourceDecompressor : IResourceDecompressor
    {
        private readonly Assembly _assembly;

        public EmbeddedResourceDecompressor()
            : this(Assembly.GetCallingAssembly())
        { }

        public EmbeddedResourceDecompressor(Assembly assembly)
            => _assembly = assembly;

        private Stream GetResourceStream(string resourceName)
        {
            var resourcePath = _assembly
                .GetManifestResourceNames()
                .FirstOrDefault(name => name.EndsWith(resourceName))    
                    ?? throw new ResourceNotFoundException(resourceName);

            return _assembly.GetManifestResourceStream(resourcePath)
                ?? throw new ResourceNotLoadedException(resourceName);
        }

        public XDocument Decompress(string resourceName)
        {
            using var stream = GetResourceStream(resourceName);
            using var zip = new GZipStream(stream, CompressionMode.Decompress);
            using var reader = new StreamReader(zip);
		
            return XDocument.Load(reader);
        }
    }
}