using System;
using System.IO;
using System.IO.Packaging;
using OfficeFlow.OpenXml.Packaging;
using OfficeFlow.OpenXml.Resources;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Interfaces;
using OfficeFlow.Word.OpenXml.Enums;

namespace OfficeFlow.Word.OpenXml;

public sealed class OpenXmlWordProcessor : IWordProcessor
{
    public static OpenXmlWordProcessor Create(OpenXmlWordDocumentType documentType)
    {
        var package = OpenXmlPackage.Create();

        var content = new EmbeddedResourceDecompressor()
            .Decompress("wordDocument.xml.gz");

        package.AddPart(
            OpenXmlPackagePart.Create(
                uri: new Uri("/word/document.xml", UriKind.Relative),
                OpenXmlWordContentTypes.GetContentType(documentType),
                CompressionOption.Normal,
                OpenXmlWordRelationships.Document,
                content));

        return new OpenXmlWordProcessor(package);
    }

    public static OpenXmlWordProcessor Open(Stream stream)
        => new(OpenXmlPackage.Open(stream));

    public static OpenXmlWordProcessor Open(string filePath)
        => new(OpenXmlPackage.Open(filePath));

    private bool _isDisposed;
    private readonly OpenXmlPackage _package;

    private Body? _body;

    private OpenXmlWordProcessor(OpenXmlPackage package)
        => _package = package;

    private OpenXmlPackagePart MainDocumentPart
    {
        get
        {
            var uri = new Uri("/word/document.xml", UriKind.Relative);

            if (_package.TryGetPart(uri, out var packagePart))
                return packagePart;

            throw new InvalidOperationException(
                "Main document part should be added to package");
        }
    }

    public OpenXmlWordDocumentType DocumentType
        => OpenXmlWordContentTypes
            .GetDocumentType(MainDocumentPart.ContentType);

    /// <inheritdoc />
    public Body Body
    {
        get
        {
            if (_body != null)
                return _body;

            var body = new Body();

            new OpenXmlElementReader(MainDocumentPart.Root)
                .Visit(body);

            return _body = body;
        }
    }

    /// <inheritdoc />
    public void Save()
        => ExecuteSave(() => _package.Save());

    /// <inheritdoc />
    public void SaveTo(Stream stream)
        => ExecuteSave(() => _package.SaveTo(stream));

    /// <inheritdoc />
    public void SaveTo(string filePath)
        => ExecuteSave(() => _package.SaveTo(filePath));

    private void ExecuteSave(Action saveStrategy)
    {
        new OpenXmlElementWriter(MainDocumentPart.Root)
            .Visit(Body);

        saveStrategy.Invoke();
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_isDisposed)
        {
            return;
        }

        _package.Dispose();

        _isDisposed = true;
    }
}