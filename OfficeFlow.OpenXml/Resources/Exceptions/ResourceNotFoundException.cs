using System;

namespace OfficeFlow.OpenXml.Resources.Exceptions
{
    public sealed class ResourceNotFoundException : Exception
    {
        public ResourceNotFoundException(string resourceName, Exception? innerException = null)
            : base($"Resource \"{resourceName}\" not found", innerException)
        { }
    }
}