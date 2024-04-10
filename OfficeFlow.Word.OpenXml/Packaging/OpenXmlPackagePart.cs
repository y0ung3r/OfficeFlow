using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;

namespace OfficeFlow.Word.OpenXml.Packaging
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
        
        public string? RelationshipType { get; }
        
        public OpenXmlPackagePart? Parent { get; private set; }

        public XElement Root
            => _xml.Root 
               ?? throw new InvalidOperationException("Package part cannot be empty");
        
        private OpenXmlPackagePart(
            Uri uri, 
            string contentType, 
            CompressionOption compressionMode, 
            string? relationshipType, 
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
            var synchronizedStream = 
                Stream.Synchronized(contentStream);
            
            using var contentWriter = 
                new StreamWriter(synchronizedStream, leaveOpen: true);
            
            _xml.Save(contentWriter, SaveOptions.None);
        }
    }
}