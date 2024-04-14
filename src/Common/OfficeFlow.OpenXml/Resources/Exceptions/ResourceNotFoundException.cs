using System;
using JetBrains.Annotations;

namespace OfficeFlow.OpenXml.Resources.Exceptions
{
    public sealed class ResourceNotFoundException : Exception
    {
        public ResourceNotFoundException(string resourceName, [CanBeNull] Exception innerException = null)
            : base($"Resource \"{resourceName}\" not found", innerException)
        { }
    }
}