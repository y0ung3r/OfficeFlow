using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;
using OfficeFlow.Word.Core.Styling;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text
{
    public sealed class RunFormat
    {
        public static RunFormat Default
            => new RunFormat
            {
                IsItalic = false,
                IsBold = false,
                IsUnderline = false,
                ForegroundColor = RgbColor.Auto,
                StrikethroughType = StrikethroughType.None
            };
        
        public bool IsItalic { get; set; }
        
        public bool IsBold { get; set; }
        
        public bool IsUnderline { get; set; }

        public RgbColor ForegroundColor { get; set; }

        public StrikethroughType StrikethroughType { get; set; }

        public TextBorder TextBorder { get; } = TextBorder.Default;
    }
}