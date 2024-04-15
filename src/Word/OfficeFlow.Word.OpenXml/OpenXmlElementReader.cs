using System.Xml.Linq;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.OpenXml;

internal sealed class OpenXmlElementReader(XElement xml) : IWordVisitor
{
    public XElement Xml { get; } = xml;

    /// <inheritdoc />
    public void Visit(Body body)
    {
        var bodyXml = Xml.Element(OpenXmlNamespaces.Word + "body");

        if (bodyXml is null)
            return;

        foreach (var bodyChild in bodyXml.Elements())
        {
            var element = bodyChild.Name.LocalName switch
            {
                "sectPr" => new Section(),
                "p" => new Paragraph(),
                _ => (Element)new UnknownElement()
            };

            if (element is IVisitable visitable)
                visitable.Accept(
                    new OpenXmlElementReader(bodyChild));

            body.AppendChild(element);
        }
    }

    /// <inheritdoc />
    public void Visit(Section section)
    { }

    /// <inheritdoc />
    public void Visit(Paragraph paragraph)
    { }

    /// <inheritdoc />
    public void Visit(Run run)
    { }

    /// <inheritdoc />
    public void Visit(LineBreak @break)
    { }

    /// <inheritdoc />
    public void Visit(HorizontalTabulation tabulation)
    { }

    /// <inheritdoc />
    public void Visit(VerticalTabulation tabulation)
    { }

    /// <inheritdoc />
    public void Visit(TextHolder text)
    { }

    /// <inheritdoc />
    public void Visit(LastRenderedPageBreak lastRenderedPageBreak)
    { }
}