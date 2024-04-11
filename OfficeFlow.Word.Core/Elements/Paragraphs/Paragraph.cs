using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs
{
    public sealed class Paragraph : CompositeElement, IVisitable
    {
        /// <inheritdoc />
        public void Accept(IWordVisitor visitor)
            => visitor.Visit(this);
    }
}