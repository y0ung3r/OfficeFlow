using System.Linq;
using FluentAssertions;
using OfficeFlow.OpenXml.Extensions;
using Xunit;

namespace OfficeFlow.OpenXml.Tests;

public sealed class StringExtensionsTests
{
    [Fact]
    public void Should_remove_restricted_xml_characters_properly()
        => StringExtensions
            .RestrictedXmlCharacters
            .Select(restrictedCharacter => $"Text without {restrictedCharacter} restricted XML character")
            .Should()
            .AllSatisfy(text =>
                text.RemoveRestrictedXmlCharacters()
                    .Should()
                    .Be("Text without  restricted XML character"));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    public void Should_do_nothing_if_text_is_null_or_empty(string? text)
        => text
            .RemoveRestrictedXmlCharacters()
            .Should()
            .Be(text);
}