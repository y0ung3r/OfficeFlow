using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Elements.Paragraphs.Text.Enums;

namespace OfficeFlow.Word.Elements.Paragraphs.Text
{
    public sealed class LineBreak : Element
    {
        public LineBreakType Type { get; set; }

        public LineBreak(LineBreakType type)
            => Type = type;
        
        public override string ToString()
            => Type switch
            {
                LineBreakType.TextWrapping => "\n",
                _ => string.Empty
            };
    }
}