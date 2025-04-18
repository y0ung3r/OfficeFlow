﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.OpenXml.Extensions;
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

internal sealed class OpenXmlElementWriter : IWordVisitor
{
    public XElement Xml { get; }
    
    public OpenXmlElementWriter(XElement xml)
    {
        Xml = xml;
        
        Xml.Elements().Remove();
    }
    
    /// <inheritdoc />
    public void Visit(Body body)
    {
        if (Xml.Name.LocalName is not "document")
            throw new InvalidOperationException(
                $"Expected: \"document\" node. But was: \"{Xml.Name.LocalName}\" node");
        
        var bodyXml = new XElement(OpenXmlNamespaces.Word + "body");
        
        foreach (var section in body.OfType<Section>())
        {
            var isMainSection = section == body.LastChild;
            
            var sectionXml = new XElement(OpenXmlNamespaces.Word + "sectPr");
            
            section.Accept(
                new OpenXmlElementWriter(sectionXml));
            
            foreach (var paragraph in section.OfType<Paragraph>())
            {
                var paragraphXml = new XElement(OpenXmlNamespaces.Word + "p");
                
                paragraph.Accept(
                    new OpenXmlElementWriter(paragraphXml));

                if (!isMainSection && paragraph == section.LastChild)
                    paragraphXml.AddFirst(sectionXml);
                
                bodyXml.Add(paragraphXml);
            }
            
            if (isMainSection)
                bodyXml.Add(sectionXml);
        }
        
        Xml.Add(bodyXml);
    }

    /// <inheritdoc />
    public void Visit(Section section)
    { }

    /// <inheritdoc />
    public void Visit(Paragraph paragraph)
    {
        var paragraphFormatXml = new XElement(OpenXmlNamespaces.Word + "pPr");

        paragraph.Format.Accept(
            new OpenXmlElementWriter(paragraphFormatXml));
        
        Xml.Add(paragraphFormatXml);

        foreach (var child in paragraph)
        {
            var childName = child switch
            {
                Run => OpenXmlNamespaces.Word + "r",
                _ => null
            };

            if (childName is null)
                return;
            
            var childXml = new XElement(childName);
            
            if (child is IVisitable visitable)
                visitable.Accept(
                    new OpenXmlElementWriter(childXml));
            
            Xml.Add(childXml);
        }
    }

    /// <inheritdoc />
    public void Visit(ParagraphFormat paragraphFormat)
    {
        VisitHorizontalAlignment(paragraphFormat);
        VisitTextAlignment(paragraphFormat);
        VisitSpacing(paragraphFormat);
        VisitKeepLines(paragraphFormat);
        VisitKeepNext(paragraphFormat);
    }

    private void VisitHorizontalAlignment(ParagraphFormat paragraphFormat)
    {
        if (paragraphFormat.HorizontalAlignment is HorizontalAlignment.Left)
            return;

        var value = paragraphFormat.HorizontalAlignment switch
        {
            // TODO[#13]: Add support for different versions of the Open XML
            HorizontalAlignment.Right => "end", // or "right"
            HorizontalAlignment.Center => "center",
            HorizontalAlignment.Both => "both",
            HorizontalAlignment.Distribute => "distribute",
            _ => null
        };
 
        if (value is null)
            return;
        
        var horizontalAlignmentXml = new XElement(
            OpenXmlNamespaces.Word + "jc",
            new XAttribute(OpenXmlNamespaces.Word + "val", value));
            
        Xml.Add(horizontalAlignmentXml);
    }

    private void VisitTextAlignment(ParagraphFormat paragraphFormat)
    {
        if (paragraphFormat.TextAlignment is TextAlignment.Auto)
            return;
        
        var value = paragraphFormat.TextAlignment switch
        {
            TextAlignment.Baseline => "baseline",
            TextAlignment.Bottom => "bottom",
            TextAlignment.Center => "center",
            TextAlignment.Top => "top",
            _ => null
        };
        
        if (value is null)
            return;
        
        var textAlignmentXml = new XElement(
            OpenXmlNamespaces.Word + "textAlignment",
            new XAttribute(OpenXmlNamespaces.Word + "val", value));
            
        Xml.Add(textAlignmentXml);
    }

    private void VisitSpacing(ParagraphFormat paragraphFormat)
    {
        var spacingProperties = new Dictionary<string, Func<IParagraphSpacing>>
        {
            { "before", () => paragraphFormat.SpacingBefore },
            { "after", () => paragraphFormat.SpacingAfter }
        };
        
        var spacingXml = new XElement(OpenXmlNamespaces.Word + "spacing");

        foreach (var spacingProperty in spacingProperties)
        {
            var valueXml = spacingProperty.Value.Invoke() switch
            {
                AutoSpacing => new XAttribute(
                    OpenXmlNamespaces.Word + spacingProperty.Key + "Autospacing", 
                    value: "true"),
                
                ExactSpacing exactSpacing => new XAttribute(
                    OpenXmlNamespaces.Word + spacingProperty.Key, 
                    exactSpacing.Value.To(AbsoluteUnits.Twips).Raw),
                
                _ => null
            };
            
            if (valueXml is null)
                continue;
            
            spacingXml.Add(valueXml);
        }
        
        var lineSpacing = paragraphFormat.SpacingBetweenLines switch
        {
            ExactSpacing exactSpacing => ("exactly", exactSpacing.Value.To(AbsoluteUnits.Twips).Raw),
            AtLeastSpacing atLeastSpacing => ("atLeast", atLeastSpacing.Value.To(AbsoluteUnits.Twips).Raw),
            MultipleSpacing multipleSpacing => ("auto", multipleSpacing.Factor * 240.0),
            _ => default((string RuleName, double Twips)?)
        };

        if (lineSpacing is not null)
        {
            var lineRuleXml = new XAttribute(
                OpenXmlNamespaces.Word + "lineRule",
                lineSpacing.Value.RuleName);
        
            var lineXml = new XAttribute(
                OpenXmlNamespaces.Word + "line", 
                lineSpacing.Value.Twips);
            
            spacingXml.Add(lineRuleXml, lineXml);
        }
        
        if (!spacingXml.HasAttributes)
            return;
        
        Xml.Add(spacingXml);
    }

    private void VisitKeepLines(ParagraphFormat paragraphFormat)
    {
        if (!paragraphFormat.KeepLines)
            return;
        
        Xml.Add(
            new XElement(OpenXmlNamespaces.Word + "keepLines"));
    }

    private void VisitKeepNext(ParagraphFormat paragraphFormat)
    {
        if (!paragraphFormat.KeepNext)
            return;
        
        Xml.Add(
            new XElement(OpenXmlNamespaces.Word + "keepNext"));
    }
    
    /// <inheritdoc />
    public void Visit(Run run)
    {
        var runFormatXml = new XElement(OpenXmlNamespaces.Word + "rPr");
        
        run.Format.Accept(
            new OpenXmlElementWriter(runFormatXml));
        
        Xml.Add(runFormatXml);

        foreach (var child in run)
        {
            var childName = child switch
            {
                LineBreak => OpenXmlNamespaces.Word + "br",
                TextHolder => OpenXmlNamespaces.Word + "t",
                VerticalTabulation => OpenXmlNamespaces.Word + "ptab",
                HorizontalTabulation => OpenXmlNamespaces.Word + "tab",
                _ => null
            };

            if (childName is null)
                return;
            
            var childXml = new XElement(childName);
            
            if (child is IVisitable visitable)
                visitable.Accept(
                    new OpenXmlElementWriter(childXml));
            
            Xml.Add(childXml);
        }
    }

    /// <inheritdoc />
    public void Visit(RunFormat runFormat)
    {
        if (runFormat.IsItalic)
        {
            var italicXml = new XElement(OpenXmlNamespaces.Word + "i");
            
            Xml.Add(italicXml);
        }

        if (runFormat.IsBold)
        {
            var boldXml = new XElement(OpenXmlNamespaces.Word + "b");
            
            Xml.Add(boldXml);
        }

        if (runFormat.IsHidden)
        {
            var vanishXml = new XElement(OpenXmlNamespaces.Word + "vanish");
            
            Xml.Add(vanishXml);
        }
        
        if (runFormat.StrikethroughType != StrikethroughType.None)
        {
            var strikethroughName = runFormat.StrikethroughType switch
            {
                StrikethroughType.Single => OpenXmlNamespaces.Word + "strike",
                StrikethroughType.Double => OpenXmlNamespaces.Word + "dstrike",
                _ => null
            };
            
            if (strikethroughName is null)
                return;
            
            Xml.Add(
                new XElement(strikethroughName));
        }
    }

    /// <inheritdoc />
    public void Visit(LineBreak @break)
    {
        if (@break.Type is LineBreakType.TextWrapping)
            return;
        
        var breakType = @break.Type switch
        {
            LineBreakType.Column => "column",
            LineBreakType.Page => "page",
            _ => null
        };

        if (breakType is null)
            return;

        Xml.Add(
            new XAttribute(OpenXmlNamespaces.Word + "type", breakType));
    }

    /// <inheritdoc />
    public void Visit(TextHolder text)
    {
        var value = text
            .Value
            .RemoveRestrictedXmlCharacters();
        
        if (string.IsNullOrEmpty(value))
            return;
        
        if (value.Any(char.IsWhiteSpace))
            Xml.Add(
                new XAttribute(XNamespace.Xml + "space", "preserve"));
        
        Xml.Value = value;
    }
}