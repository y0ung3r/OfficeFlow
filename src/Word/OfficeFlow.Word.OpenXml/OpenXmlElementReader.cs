﻿using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Enums;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;
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
        var horizontalAlignment =
            xml.Element(OpenXmlNamespaces.Word + "jc")?
                .Attribute(OpenXmlNamespaces.Word + "val")?
                .Value;

        paragraphFormat.HorizontalAlignment = horizontalAlignment switch
        {
            // TODO: Add support for different versions of the Open XML
            "start" or "left" => HorizontalAlignment.Left,
            "end" or "right" => HorizontalAlignment.Right,
            "center" => HorizontalAlignment.Center,
            "both" => HorizontalAlignment.Both,
            "distribute" => HorizontalAlignment.Distribute,
            _ => ParagraphFormat.Default.HorizontalAlignment
        };
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
                
                case "u":
                    // TODO: Add underline support
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