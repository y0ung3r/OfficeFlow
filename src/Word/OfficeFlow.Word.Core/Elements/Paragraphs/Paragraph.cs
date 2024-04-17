using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs;

public sealed class Paragraph(ParagraphFormat format) : CompositeElement, IVisitable
{
    public ParagraphFormat Format { get; } = format;

    public string Text
        => string.Concat(
            this.Where(child => child is Run));
    
    public Paragraph()
        : this(ParagraphFormat.Default)
    { }

    public void AppendText()
        => AppendText("\n");

    public void AppendText(string text)
    {
        var runs = new List<Run>();
        var stringBuilder = new StringBuilder();
        var delimeters = new[] { '\t', '\n', '\r' };
		
        for (var index = 0; index < text.Length; index++)
        {
            var symbol = text[index];
            var builderHasText = stringBuilder.Length > 0;
            var isLastSymbol = index == text.Length - 1;
            var isDelimeterSymbol = delimeters.Contains(symbol);

            if (!isDelimeterSymbol)
            {
                stringBuilder.Append(symbol);
            }
			
            if (builderHasText && (isDelimeterSymbol || isLastSymbol))
            {
                var run = new Run();
                
                var line = stringBuilder.ToString();
                run.AppendText(line);
                
                runs.Add(run);
                
                stringBuilder.Clear();
            }

            if (isDelimeterSymbol)
            {
                var previousRun = runs.LastOrDefault();
                var delimeterRun = new Run();

                switch (symbol)
                {
                    case '\t':
                        delimeterRun.AppendTabulation();
                        runs.Add(delimeterRun);
                        break;
					
                    case '\r':
                    case '\n':
                        if (previousRun is not null && !previousRun.HasLineBreaks)
                        {
                            delimeterRun.AppendLineBreak();
                            runs.Add(delimeterRun);
                        }
                        break;
                }
            }
        }

        foreach (var run in runs)
        {
            AppendChild(run);
        }
    }
    
    /// <inheritdoc />
    public void Accept(IWordVisitor visitor)
        => visitor.Visit(this);

    /// <inheritdoc />
    public override string ToString()
        => Text;
}