using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;

namespace OfficeFlow.Word.Core.Interfaces;

public interface IWordVisitor
{
    /// <summary>
    /// Visit <see cref="Body"/>
    /// </summary>
    /// <param name="body">Instance of <see cref="Body"/></param>
    void Visit(Body body);

    /// <summary>
    /// Visit <see cref="Section"/>
    /// </summary>
    /// <param name="section">Instance of <see cref="Section"/></param>
    void Visit(Section section);

    /// <summary>
    /// Visit <see cref="Paragraph"/>
    /// </summary>
    /// <param name="paragraph">Instance of <see cref="Paragraph"/></param>
    void Visit(Paragraph paragraph);

    /// <summary>
    /// Visit <see cref="ParagraphFormat"/>
    /// </summary>
    /// <param name="paragraphFormat">Instance of <see cref="ParagraphFormat"/></param>
    void Visit(ParagraphFormat paragraphFormat);

    /// <summary>
    /// Visit <see cref="Run"/>
    /// </summary>
    /// <param name="run">Instance of <see cref="Run"/></param>
    void Visit(Run run);

    /// <summary>
    /// Visit <see cref="RunFormat"/>
    /// </summary>
    /// <param name="runFormat">Instance of <see cref="RunFormat"/></param>
    void Visit(RunFormat runFormat);

    /// <summary>
    /// Visit <see cref="LineBreak"/>
    /// </summary>
    /// <param name="break">Instance of <see cref="LineBreak"/></param>
    void Visit(LineBreak @break);

    /// <summary>
    /// Visit <see cref="TextHolder"/>
    /// </summary>
    /// <param name="text">Instance of <see cref="TextHolder"/></param>
    void Visit(TextHolder text);
}