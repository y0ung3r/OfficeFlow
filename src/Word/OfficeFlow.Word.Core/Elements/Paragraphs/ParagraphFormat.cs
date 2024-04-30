﻿using OfficeFlow.MeasureUnits.Absolute;
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

    public TextAlignment TextAlignment { get; set; }
        = TextAlignment.Auto;
    
    public IParagraphSpacing SpacingBefore { get; set; } 
        = ParagraphSpacing.Exactly(0.0, AbsoluteUnits.Points);

    public IParagraphSpacing SpacingAfter { get; set; } 
        = ParagraphSpacing.Exactly(8.0, AbsoluteUnits.Points);
    
    public bool KeepLines { get; set; }
    
    public bool KeepNext { get; set; }

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}