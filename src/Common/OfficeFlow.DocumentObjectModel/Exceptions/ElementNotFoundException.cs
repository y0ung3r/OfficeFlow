using System;

namespace OfficeFlow.DocumentObjectModel.Exceptions
{
    public sealed class ElementNotFoundException : Exception
    {
        public ElementNotFoundException(string message)
            : base(message)
        { }
    }
}