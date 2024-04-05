using OfficeFlow.Word.Elements.Enums;

namespace OfficeFlow.Word.Elements.Paragraphs.Text
{
    public sealed class TextBorder
    {
        public static TextBorder Default = new TextBorder
        {
            Type = BorderType.None,
            UseFrameEffect = false,
            UseShadowEffect = false
        };
        
        public BorderType Type { get; set; }
        
        public bool UseFrameEffect { get; set; }
        
        public bool UseShadowEffect { get; set; }
    }
}