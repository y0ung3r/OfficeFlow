using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting.Enums;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting;

public sealed class RunFormat : IVisitable
{
    public static RunFormat Default
        => new();

    public bool IsItalic { get; set; }

    public bool IsBold { get; set; }
    
    public bool IsHidden { get; set; }

    public StrikethroughType StrikethroughType { get; set; }
        = StrikethroughType.None;

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}