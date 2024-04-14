using System.Linq;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Text
{
    public sealed class Run : CompositeElement
    {
        public RunFormat Format { get; }
        
        public string Text 
            => string.Concat(
                this.Where(child => 
                    child is LineBreak || child is TextHolder || child is HorizontalTabulation));

        public bool HasLineBreaks 
            => this
                .OfType<LineBreak>()
                .Any(child => child.Type is LineBreakType.TextWrapping);

        public bool IsLastOnPage
            => this.Any(child =>
                child is LastRenderedPageBreak);
        
        public Run()
            : this(RunFormat.Default)
        { }

        public Run(RunFormat format)
            => Format = format;

        public void AppendTabulation()
            => AppendChild(new HorizontalTabulation());

        public void AppendLineBreak()
            => AppendChild(new LineBreak(LineBreakType.TextWrapping));

        public void AppendText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            AppendChild(new TextHolder(value));
        }

        public override string ToString()
            => Text;
    }
}