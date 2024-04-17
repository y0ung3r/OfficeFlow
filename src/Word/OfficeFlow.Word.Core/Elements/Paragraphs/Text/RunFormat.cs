using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text;

public sealed class RunFormat : IVisitable
{
    public static RunFormat Default
        => new()
        {
            IsItalic = false,
            IsBold = false,
            StrikethroughType = StrikethroughType.None
        };

    public bool IsItalic { get; set; }

    public bool IsBold { get; set; }

    public StrikethroughType StrikethroughType { get; set; }

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}