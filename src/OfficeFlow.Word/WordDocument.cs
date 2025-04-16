using System;
using System.Collections.Generic;
using System.IO;
using OfficeFlow.Formats;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Enums;
using OfficeFlow.Word.Core.Interfaces;
using OfficeFlow.Word.Extensions;
using OfficeFlow.Word.OpenXml;

namespace OfficeFlow.Word;

public sealed class WordDocument : IDisposable
{
    private bool _isDisposed;
    private readonly DocumentSettings _settings;
    private readonly IWordProcessor _processor;

    public WordDocument(WordDocumentType documentType)
        : this(documentType, DocumentSettings.Default)
    { }

    public WordDocument(WordDocumentType documentType, DocumentSettings settings)
        : this(settings, CreateProcessor(documentType))
    { }

    public WordDocument(Stream stream)
        : this(stream, DocumentSettings.Default)
    { }

    public WordDocument(Stream stream, DocumentSettings settings)
        : this(settings, OpenProcessor(stream))
    { }

    public WordDocument(string filePath)
        : this(filePath, DocumentSettings.Default)
    { }

    public WordDocument(string filePath, DocumentSettings settings)
        : this(settings, OpenProcessor(filePath))
    { }

    private WordDocument(DocumentSettings settings, IWordProcessor processor)
    {
        _processor = processor;
        _settings = settings;
    }

    public IEnumerable<Section> Sections
        => _processor.Body.Sections;

    public Section LastSection
        => _processor.Body.LastSection;

    public Section AppendSection()
        => _processor.Body.AppendSection();

    public void Save()
        => _processor.Save();

    public void SaveTo(Stream stream)
        => _processor.SaveTo(stream);

    public void SaveTo(string filePath)
        => _processor.SaveTo(filePath);

    public void Close()
        => Dispose();

    /// <inheritdoc />
    public void Dispose()
    {
        if (_isDisposed)
        {
            return;
        }

        _processor.Dispose();

        _isDisposed = true;
    }

    private static IWordProcessor CreateProcessor(WordDocumentType documentType)
    {
        var format = documentType.ToOfficeFormat();

        if (format is OpenXmlFormat)
            return OpenXmlWordProcessor.Create(
                documentType.ToOpenXmlType());

        if (format is BinaryFormat)
            throw new NotSupportedException();

        throw new NotSupportedException();
    }

    private static IWordProcessor OpenProcessor(Stream stream)
    {
        var format = OfficeFormatDetector.Detect(stream);

        if (format is OpenXmlFormat)
            return OpenXmlWordProcessor.Open(stream);

        if (format is BinaryFormat)
            throw new NotSupportedException();

        throw new NotSupportedException();
    }

    private static IWordProcessor OpenProcessor(string filePath)
    {
        var format = OfficeFormatDetector.Detect(filePath);

        if (format is OpenXmlFormat)
            return OpenXmlWordProcessor.Open(filePath);

        if (format is BinaryFormat)
            throw new NotSupportedException();

        throw new NotSupportedException();
    }
}