namespace OfficeFlow.Word.Core.Interfaces;

public interface IVisitable
{
    /// <summary>
    /// Accept <see cref="IWordVisitor"/>
    /// </summary>
    /// <param name="visitor">Instance of <see cref="IWordVisitor"/></param>
    void Accept(IWordVisitor visitor);
}