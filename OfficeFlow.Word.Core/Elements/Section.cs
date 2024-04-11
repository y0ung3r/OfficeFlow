using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements
{
    public class Section : CompositeElement, IVisitable
    {
        /// <inheritdoc />
        public void Accept(IWordVisitor visitor)
            => visitor.Visit(this);
    }
}