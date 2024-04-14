using System;
using FluentAssertions;
using OfficeFlow.Word.Core.Enums;
using Xunit;

namespace OfficeFlow.Word.Tests;

public sealed class WordDocumentTests
{
    [Theory]
    [InlineData(WordDocumentType.Docx)]
    [InlineData(WordDocumentType.Docm)]
    [InlineData(WordDocumentType.Dotx)]
    [InlineData(WordDocumentType.Dotm)]
    public void Should_create_word_document_properly(WordDocumentType documentType)
        // ReSharper disable once ObjectCreationAsStatement
        => new Action(() => new WordDocument(documentType))
            .Should()
            .NotThrow();
}