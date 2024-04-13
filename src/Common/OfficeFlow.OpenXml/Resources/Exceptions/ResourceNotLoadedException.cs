using System;

namespace OfficeFlow.OpenXml.Resources.Exceptions
{
    public sealed class ResourceNotLoadedException : Exception
    {
        public ResourceNotLoadedException(string resourceName, Exception? innerException = null)
            : base($"Failed to load resource \"{resourceName}\"", innerException)
        { }
    }
}