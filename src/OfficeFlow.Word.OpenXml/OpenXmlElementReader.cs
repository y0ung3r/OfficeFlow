using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.MeasureUnits.Extensions;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Enums;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing.Interfaces;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting.Enums;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.OpenXml;

internal sealed class OpenXmlElementReader(XElement xml) : IWordVisitor
{
    /// <inheritdoc />
    public void Visit(Body body)
    {
        var bodyXml = xml.Element(OpenXmlNamespaces.Word + "body");

        if (bodyXml is null)
        {
            body.AppendChild(
                new Section());
            
            return;
        }

        var paragraphs = new Queue<Paragraph>();
        
        foreach (var paragraphXml in bodyXml.Elements(OpenXmlNamespaces.Word + "p"))
        {
            var paragraph = new Paragraph();
                    
            paragraph.Accept(
                new OpenXmlElementReader(paragraphXml));

            paragraphs.Enqueue(paragraph);
            
            var sectionXml = paragraphXml.Element(OpenXmlNamespaces.Word + "sectPr");

            if (sectionXml is null)
                continue;

            var section = new Section();
                    
            section.Accept(
                new OpenXmlElementReader(sectionXml));

            while (paragraphs.Count > 0)
                section.AppendChild(
                    paragraphs.Dequeue());
            
            body.AppendChild(section);
        }

        var lastSection = new Section();
        
        var lastSectionXml = bodyXml
            .Elements(OpenXmlNamespaces.Word + "sectPr")
            .LastOrDefault();
        
        if (lastSectionXml is not null)
            lastSection.Accept(
                new OpenXmlElementReader(lastSectionXml));
        
        while (paragraphs.Count > 0)
            lastSection.AppendChild(
                paragraphs.Dequeue());
        
        body.AppendChild(lastSection);
    }

    /// <inheritdoc />
    public void Visit(Section section)
    { }

    /// <inheritdoc />
    public void Visit(Paragraph paragraph)
    {
        var propertiesXml = xml.Element(OpenXmlNamespaces.Word + "pPr");

        if (propertiesXml is not null)
            paragraph.Format.Accept(
                new OpenXmlElementReader(propertiesXml));

        foreach (var runXml in xml.Elements(OpenXmlNamespaces.Word + "r"))
        {
            var run = new Run();
            
            run.Accept(
                new OpenXmlElementReader(runXml));
            
            paragraph.AppendChild(run);
        }
    }

    /// <inheritdoc />
    public void Visit(ParagraphFormat paragraphFormat)
    {
        VisitHorizontalAlignment(paragraphFormat);
        VisitTextAlignment(paragraphFormat);
        VisitParagraphSpacing(paragraphFormat);
        VisitLineSpacing(paragraphFormat);
        VisitKeepLines(paragraphFormat);
        VisitKeepNext(paragraphFormat);
    }

    private void VisitHorizontalAlignment(ParagraphFormat paragraphFormat)
    {
        var alignmentXml = xml.Element(OpenXmlNamespaces.Word + "jc")?
            .Attribute(OpenXmlNamespaces.Word + "val")?
            .Value;

        var horizontalAlignment = alignmentXml switch
        {
            // TODO[#13]: Add support for different versions of the Open XML
            "start" or "left" => HorizontalAlignment.Left,
            "end" or "right" => HorizontalAlignment.Right,
            "center" => HorizontalAlignment.Center,
            "both" => HorizontalAlignment.Both,
            "distribute" => HorizontalAlignment.Distribute,
            _ => default(HorizontalAlignment?)
        };
        
        if (horizontalAlignment is null)
            return;

        paragraphFormat.HorizontalAlignment = horizontalAlignment.Value;
    }

    private void VisitTextAlignment(ParagraphFormat paragraphFormat)
    {
        var alignmentXml = xml.Element(OpenXmlNamespaces.Word + "textAlignment")?
            .Attribute(OpenXmlNamespaces.Word + "val")?
            .Value;
        
        var textAlignment = alignmentXml switch
        {
            "auto" => TextAlignment.Auto,
            "baseline" => TextAlignment.Baseline,
            "bottom" => TextAlignment.Bottom,
            "center" => TextAlignment.Center,
            "top" => TextAlignment.Top,
            _ => default(TextAlignment?)
        };
        
        if (textAlignment is null)
            return;

        paragraphFormat.TextAlignment = textAlignment.Value;
    }

    private void VisitParagraphSpacing(ParagraphFormat paragraphFormat)
    {
        var spacingProperties = new Dictionary<string, Func<IParagraphSpacing, IParagraphSpacing>>
        {
            { "before", value => paragraphFormat.SpacingBefore = value },
            { "after", value => paragraphFormat.SpacingAfter = value }
        };
        
        var spacingXml = xml.Element(OpenXmlNamespaces.Word + "spacing");
        
        if (spacingXml is null)
            return;

        foreach (var spacingProperty in spacingProperties)
        {
            var autospacingXml = spacingXml
                .Attribute(OpenXmlNamespaces.Word + spacingProperty.Key + "Autospacing");

            if (autospacingXml is { Value: "1" or "true" })
            {
                spacingProperty.Value.Invoke(ParagraphSpacing.Auto);
            
                continue;
            }

            var exactSpacingXml = spacingXml
                .Attribute(OpenXmlNamespaces.Word + spacingProperty.Key);

            if (exactSpacingXml is null)
                continue;

            spacingProperty.Value.Invoke(
                ParagraphSpacing.Exactly(
                    exactSpacingXml.Value.ToUnits(AbsoluteUnits.Twips)));
        }
    }

    private void VisitLineSpacing(ParagraphFormat paragraphFormat)
    {
        var spacingXml = xml.Element(OpenXmlNamespaces.Word + "spacing");
        
        if (spacingXml is null)
            return;
        
        var lineRuleXml = spacingXml
            .Attribute(OpenXmlNamespaces.Word + "lineRule");
        
        var lineXml = spacingXml
            .Attribute(OpenXmlNamespaces.Word + "line");
        
        if (lineRuleXml is null || lineXml is null)
            return;

        var twips = lineXml
            .Value
            .ToUnits(AbsoluteUnits.Twips);

        var spacingBetweenLines = lineRuleXml.Value switch
        {
            "auto" => LineSpacing.Multiple(twips.Raw / 240.0),
            "exactly" => LineSpacing.Exactly(twips.Raw, AbsoluteUnits.Twips),
            "atLeast" => LineSpacing.AtLeast(twips.Raw, AbsoluteUnits.Twips),
            _ => null
        };
        
        if (spacingBetweenLines is null)
            return;

        paragraphFormat.SpacingBetweenLines = spacingBetweenLines;
    }
    
    private void VisitKeepLines(ParagraphFormat paragraphFormat)
    {
        var keepLinesXml = 
            xml.Element(OpenXmlNamespaces.Word + "keepLines");
        
        paragraphFormat.KeepLines = keepLinesXml is not null;
    }

    private void VisitKeepNext(ParagraphFormat paragraphFormat)
    {
        var keepNextXml = 
            xml.Element(OpenXmlNamespaces.Word + "keepNext");
        
        paragraphFormat.KeepNext = keepNextXml is not null;
    }

    /// <inheritdoc />
    public void Visit(Run run)
    {
        var propertiesXml = xml.Element(OpenXmlNamespaces.Word + "rPr");
        
        if (propertiesXml is not null)
            run.Format.Accept(
                new OpenXmlElementReader(propertiesXml));
        
        foreach (var childXml in xml.Elements().Where(childXml => childXml.Name.LocalName != "rPr"))
        {
            var child = childXml.Name.LocalName switch
            {
                "br" => new LineBreak(),
                "delText" => new UnknownElement(),
                "t" => new TextHolder(),
                "ptab" => new VerticalTabulation(), // "\v"
                "tab" => new HorizontalTabulation(),
                "tc" => new UnknownElement(),
                "tr" => new UnknownElement(), // Line Break?
                _ => (Element)new UnknownElement()
            };
            
            if (child is IVisitable visitable)
                visitable.Accept(
                    new OpenXmlElementReader(childXml));
            
            run.AppendChild(child);
        }
    }

    /// <inheritdoc />
    public void Visit(RunFormat runFormat)
    {
        foreach (var childXml in xml.Elements())
        {
            switch (childXml.Name.LocalName)
            {
                case "i":
                    runFormat.IsItalic = true;
                    break;
                
                case "b":
                    runFormat.IsBold = true;
                    break;
                
                case "vanish":
                    runFormat.IsHidden = true;
                    break;
                
                case "u":
                    // TODO[#7]: Add underline support
                    break;
                
                case "strike":
                    runFormat.StrikethroughType = StrikethroughType.Single;
                    break;
                
                case "dstrike":
                    runFormat.StrikethroughType = StrikethroughType.Double;
                    break;
            }
        }
    }

    /// <inheritdoc />
    public void Visit(LineBreak @break)
    {
        var typeXml = xml.Attribute(OpenXmlNamespaces.Word + "type");

        if (typeXml is null)
            return;
        
        @break.Type = typeXml.Value switch
        {
            "textWrapping" => LineBreakType.TextWrapping,
            "column" => LineBreakType.Column,
            "page" => LineBreakType.Page,
            _ => LineBreakType.TextWrapping
        };
    }

    /// <inheritdoc />
    public void Visit(TextHolder text)
        => text.Value = xml.Value;
}