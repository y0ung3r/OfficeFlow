using System.Linq;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting.Enums;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text;

public sealed class Run(RunFormat format) : CompositeElement, IVisitable
{
    public RunFormat Format { get; } = format;

    public string Text
        => string.Concat(
            this.Where(child =>
                child is LineBreak or TextHolder or HorizontalTabulation));

    public bool HasLineBreaks
        => this.Any(child =>
            child is LineBreak { Type: LineBreakType.TextWrapping });

    public bool IsLastOnPage
        => this.Any(child =>
            child is LastRenderedPageBreak);

    public Run()
        : this(RunFormat.Default)
    { }

    public void AppendTabulation()
        => AppendChild(
            new HorizontalTabulation());

    public void AppendLineBreak()
        => AppendChild(
            new LineBreak(LineBreakType.TextWrapping));

    public void AppendText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        AppendChild(
            new TextHolder(text));
    }

    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);

    public override string ToString()
        => Text;
}