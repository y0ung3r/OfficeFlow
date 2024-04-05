using System;

namespace OfficeFlow.DocumentObjectModel.Exceptions
{
    public sealed class ElementHasNoParentException : Exception
    {
        public ElementHasNoParentException()
            : base("Element does not have a parent")
        { }
    }
}