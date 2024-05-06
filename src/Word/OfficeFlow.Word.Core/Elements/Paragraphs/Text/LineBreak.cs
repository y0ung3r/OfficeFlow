using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting.Enums;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text;

public sealed class LineBreak(LineBreakType type) : Element, IVisitable
{
    public LineBreakType Type { get; set; } = type;
    
    public LineBreak()
        : this(LineBreakType.TextWrapping)
    { }

    public override string ToString()
        => Type switch
        {
            LineBreakType.TextWrapping => "\n",
            _ => string.Empty
        };

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}