using System.Xml.Linq;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.OpenXml
{
    internal sealed class OpenXmlElementWriter : IWordVisitor
    {
        public XElement Xml { get; }
        
        public OpenXmlElementWriter(XElement xml)
            => Xml = xml;
        
        /// <inheritdoc />
        public void Visit(Body body)
        {
            var bodyXml = new XElement(OpenXmlNamespaces.Word + "body");

            foreach (var child in body)
            {
                XName childName = null;
                
                switch (child)
                {
                    case Section _:
                        childName = OpenXmlNamespaces.Word + "sectPr";
                        break;
                    
                    case Paragraph _:
                        childName = OpenXmlNamespaces.Word + "p";
                        break;
                }

                if (childName is null)
                    continue;

                var childXml = new XElement(childName);
                
                if (child is IVisitable visitable)
                    visitable.Accept(
                        new OpenXmlElementWriter(childXml));
                
                bodyXml.Add(childXml);
            }
            
            Xml.Add(bodyXml);
        }

        /// <inheritdoc />
        public void Visit(Section section)
        {
            
        }

        /// <inheritdoc />
        public void Visit(Paragraph paragraph)
        {
            
        }

        /// <inheritdoc />
        public void Visit(Run run)
        {
            
        }

        /// <inheritdoc />
        public void Visit(LineBreak @break)
        {
            
        }

        /// <inheritdoc />
        public void Visit(HorizontalTabulation tabulation)
        {
            
        }

        /// <inheritdoc />
        public void Visit(VerticalTabulation tabulation)
        {
            
        }

        /// <inheritdoc />
        public void Visit(TextHolder text)
        {
            
        }

        /// <inheritdoc />
        public void Visit(LastRenderedPageBreak lastRenderedPageBreak)
        {
            
        }
    }
}