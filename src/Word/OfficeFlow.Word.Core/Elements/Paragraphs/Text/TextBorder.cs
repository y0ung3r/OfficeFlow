using OfficeFlow.Word.Core.Elements.Enums;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text;

public sealed class TextBorder
{
    public static readonly TextBorder Default = new()
    {
        Type = BorderType.None,
        UseFrameEffect = false,
        UseShadowEffect = false
    };

    public BorderType Type { get; set; }

    public bool UseFrameEffect { get; set; }

    public bool UseShadowEffect { get; set; }
}