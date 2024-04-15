using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text;

public sealed class LineBreak : Element
{
    public LineBreakType Type { get; set; }

    public LineBreak(LineBreakType type)
        => Type = type;

    public override string ToString()
    {
        if (Type == LineBreakType.TextWrapping)
            return "\n";

        return string.Empty;
    }
}