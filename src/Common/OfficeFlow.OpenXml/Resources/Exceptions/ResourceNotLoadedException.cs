using System;
using JetBrains.Annotations;

namespace OfficeFlow.OpenXml.Resources.Exceptions
{
    public sealed class ResourceNotLoadedException : Exception
    {
        public ResourceNotLoadedException(string resourceName, [CanBeNull] Exception innerException = null)
            : base($"Failed to load resource \"{resourceName}\"", innerException)
        { }
    }
}