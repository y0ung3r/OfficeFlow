using System.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.Core.Tests;

public sealed class ParagraphTests
{
    [Fact]
    public void Should_create_default_format_for_paragraph()
    {
        // Arrange
        var sut = new Paragraph();
        
        // Act & Assert
        sut.Format
            .Should()
            .BeEquivalentTo(ParagraphFormat.Default);
    }

    [Theory]
    [InlineData("\n\ttext\ntext     \r\n     text")]
    [InlineData("text\ntext")]
    public void Should_return_text_properly(string expectedText)
    {
        // Arrange
        var sut = new Paragraph();
        
        sut.AppendText(expectedText);
        
        // Act & Assert
        sut.Text
            .Should()
            .Be(expectedText);
    }
    
    [Fact]
    public void Should_append_text_properly()
    {
        // Arrange
        var sut = new Paragraph();

        // Act
        sut.AppendText("\n\ttext\ntext     \r\n     text");
        
        // Assert
        sut.OfType<Run>()
            .Should()
            .SatisfyRespectively(
                run => run
                    .Should()
                    .ContainItemsAssignableTo<LineBreak>(),
                run => run
                    .Should()
                    .ContainItemsAssignableTo<HorizontalTabulation>(),
                run => run
                    .OfType<TextHolder>()
                    .Should()
                    .SatisfyRespectively(textHolder =>
                        textHolder
                            .Value
                            .Should()
                            .Be("text")),
                run => run
                    .Should()
                    .ContainItemsAssignableTo<LineBreak>(),
                run => run
                    .OfType<TextHolder>()
                    .Should()
                    .SatisfyRespectively(textHolder =>
                        textHolder
                            .Value
                            .Should()
                            .Be("text     \r")),
                run => run
                    .Should()
                    .ContainItemsAssignableTo<LineBreak>(),
                run => run
                    .OfType<TextHolder>()
                    .Should()
                    .SatisfyRespectively(textHolder =>
                        textHolder
                            .Value
                            .Should()
                            .Be("     text")));
    }
    
    [Fact]
    public void Should_append_line_break()
    {
        // Arrange
        var sut = new Paragraph();

        // Act
        sut.AppendText();
        
        // Assert
        sut.OfType<Run>()
            .Should()
            .SatisfyRespectively(run =>
                run.Should()
                    .ContainItemsAssignableTo<LineBreak>());
        
        sut.Text
            .Should()
            .Be("\n");
    }

    [Fact]
    public void Should_append_horizontal_tabulation()
    {
        // Arrange
        var sut = new Paragraph();
        
        // Act
        sut.AppendText("\t");
        
        // Assert
        sut.OfType<Run>()
            .Should()
            .SatisfyRespectively(run =>
                run.Should()
                    .ContainItemsAssignableTo<HorizontalTabulation>());
        
        sut.Text
            .Should()
            .Be("\t");
    }
}