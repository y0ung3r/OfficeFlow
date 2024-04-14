using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace OfficeFlow.OpenXml.Packaging
{
    public sealed class OpenXmlPackagePart
    {
        public static OpenXmlPackagePart Create(
            Uri uri, 
            string contentType, 
            CompressionOption compressionMode, 
            string relationshipType, 
            XDocument content)
            => new OpenXmlPackagePart(
                uri, 
                contentType, 
                compressionMode, 
                relationshipType, 
                content);
        
        public static OpenXmlPackagePart Open(
            Uri uri, 
            string contentType, 
            CompressionOption compressionMode, 
            Stream contentStream)
        {
            var content = XDocument.Load(contentStream);
            
            return new OpenXmlPackagePart(
                uri, 
                contentType, 
                compressionMode, 
                relationshipType: null, 
                content);
        }
        
        private readonly XDocument _xml;
        
        public Uri Uri { get; }

        public string ContentType { get; }
        
        public CompressionOption CompressionMode { get; }
        
        [CanBeNull] 
        public string RelationshipType { get; }
        
        [CanBeNull] 
        public OpenXmlPackagePart Parent { get; private set; }

        public XElement Root
            => _xml.Root 
               ?? throw new InvalidOperationException("Package part cannot be empty");
        
        private OpenXmlPackagePart(
            Uri uri, 
            string contentType, 
            CompressionOption compressionMode, 
            [CanBeNull] string relationshipType, 
            XDocument xml)
        {
            Uri = uri;
            ContentType = contentType;
            CompressionMode = compressionMode;
            RelationshipType = relationshipType;
            
            _xml = xml;
        }

        public void AddChild(OpenXmlPackagePart childPart)
        {
            childPart.Parent = this;
        }
        
        public void FlushTo(Stream contentStream)
        {
            var synchronizedStream = Stream.Synchronized(contentStream);
            var contentWriter = new StreamWriter(synchronizedStream);
            
            _xml.Save(contentWriter, SaveOptions.None);
            
            contentWriter.Flush();
        }
    }
}