using OfficeFlow.Word.Core.Elements.Paragraphs.Enums;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs;

public sealed record ParagraphFormat : IVisitable
{
    public static ParagraphFormat Default
        => new()
        {
            HorizontalAlignment = HorizontalAlignment.Left
        };
    
    public HorizontalAlignment HorizontalAlignment { get; set; }

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}