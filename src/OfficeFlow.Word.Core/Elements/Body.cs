using System.Collections.Generic;
using System.Linq;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements;

public sealed class Body : CompositeElement, IVisitable
{
    public IEnumerable<Section> Sections
        => this.OfType<Section>();
    
    public Section LastSection
    {
        get
        {
            if (LastChild is Section lastSection)
                return lastSection;

            return AppendSection();
        }
    }

    public Section AppendSection()
    {
        var section = new Section();
        
        AppendChild(section);

        return section;
    }
    
    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);
}