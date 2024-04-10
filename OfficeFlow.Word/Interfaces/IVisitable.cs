namespace OfficeFlow.Word.Interfaces
{
    public interface IVisitable
    {
        /// <summary>
        /// Accept <see cref="IElementVisitor"/>
        /// </summary>
        /// <param name="visitor">Instance of <see cref="IElementVisitor"/></param>
        void Accept(IElementVisitor visitor);
    }
}