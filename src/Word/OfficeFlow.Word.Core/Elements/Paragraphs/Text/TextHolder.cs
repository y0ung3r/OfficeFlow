using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text;

public sealed class TextHolder(string value) : Element, IVisitable
{
    public string Value { get; set; } = value;
    
    public TextHolder() 
        : this(string.Empty)
    { }

    public override string ToString()
        => Value;

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}