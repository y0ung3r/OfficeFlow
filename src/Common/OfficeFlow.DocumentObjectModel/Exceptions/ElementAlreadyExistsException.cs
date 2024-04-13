using System;

namespace OfficeFlow.DocumentObjectModel.Exceptions
{
    public sealed class ElementAlreadyExistsException : Exception
    {
        public ElementAlreadyExistsException(string message)
            : base(message)
        { }
    }
}