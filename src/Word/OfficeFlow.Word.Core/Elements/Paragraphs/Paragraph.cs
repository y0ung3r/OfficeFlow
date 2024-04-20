using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs;

public sealed class Paragraph(ParagraphFormat format) : CompositeElement, IVisitable
{
    private static readonly char[] ControlCharacters
        = { '\t', '\n' };
    
    public ParagraphFormat Format { get; } = format;

    public string Text
        => string.Concat(
            this.Where(child => child is Run));
    
    public Paragraph()
        : this(ParagraphFormat.Default)
    { }

    public void AppendText()
        => AppendText(text: "\n");

    public void AppendText(string text)
    {
        var runs = new List<Run>();
        var stringBuilder = new StringBuilder();
		
        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            var isLastCharacter = index == text.Length - 1;
            var isControlCharacter = ControlCharacters.Contains(character);

            if (!isControlCharacter)
            {
                stringBuilder.Append(character);
            }
			
            if (stringBuilder.Length > 0 && (isControlCharacter || isLastCharacter))
            {
                var run = new Run();
                
                var line = stringBuilder.ToString();
                
                run.AppendText(line);
                
                runs.Add(run);
                
                stringBuilder.Clear();
            }

            if (!isControlCharacter)
            {
                continue;
            }

            var controlCharacterRun = new Run();

            switch (character)
            {
                case '\t':
                    controlCharacterRun.AppendTabulation();
                    break;

                case '\n':
                    controlCharacterRun.AppendLineBreak();
                    break;
            }

            if (controlCharacterRun.Count > 0)
            {
                runs.Add(controlCharacterRun);
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