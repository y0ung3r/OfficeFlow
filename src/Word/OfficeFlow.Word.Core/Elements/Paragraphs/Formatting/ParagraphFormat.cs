using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Enums;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing.Interfaces;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Formatting;

public sealed class ParagraphFormat : IVisitable
{
    public static ParagraphFormat Default
        => new();
    
    public HorizontalAlignment HorizontalAlignment { get; set; } 
        = HorizontalAlignment.Left;

    public TextAlignment TextAlignment { get; set; }
        = TextAlignment.Auto;
    
    public IParagraphSpacing SpacingBefore { get; set; } 
        = ParagraphSpacing.Exactly(0.0, AbsoluteUnits.Points);

    public IParagraphSpacing SpacingAfter { get; set; } 
        = ParagraphSpacing.Exactly(8.0, AbsoluteUnits.Points);
    
    public ILineSpacing SpacingBetweenLines { get; set; } 
        = LineSpacing.Multiple(factor: 1.08);
    
    public bool KeepLines { get; set; }
    
    public bool KeepNext { get; set; }

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}