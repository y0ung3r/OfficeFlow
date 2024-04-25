using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Enums;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing.Interfaces;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs;

public sealed class ParagraphFormat : IVisitable
{
    public static ParagraphFormat Default
        => new();
    
    public HorizontalAlignment HorizontalAlignment { get; set; } 
        = HorizontalAlignment.Left;

    public IParagraphSpacing SpacingBefore { get; set; } 
        = ParagraphSpacing.Exactly<Points>(0);

    public IParagraphSpacing SpacingAfter { get; set; } 
        = ParagraphSpacing.Exactly<Points>(8.0);

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}