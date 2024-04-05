using OfficeFlow.DocumentObjectModel;

namespace OfficeFlow.Word.Elements.Paragraphs.Text
{
    public sealed class TextHolder : Element
    {
        public string Value { get; set; }
    
        public TextHolder(string value)
            => Value = value;
        
        public override string ToString()
            => Value;
    }
}