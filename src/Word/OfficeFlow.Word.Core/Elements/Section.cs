using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements;

public sealed class Section : CompositeElement, IVisitable
{
    public Paragraph AppendParagraph()
    {
        var paragraph = new Paragraph();
        
        AppendChild(paragraph);

        return paragraph;
    }
    
    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}