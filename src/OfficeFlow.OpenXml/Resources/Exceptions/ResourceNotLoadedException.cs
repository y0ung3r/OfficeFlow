using System;

namespace OfficeFlow.OpenXml.Resources.Exceptions;

public sealed class ResourceNotLoadedException(string resourceName, Exception? innerException = null)
    : Exception($"Failed to load resource \"{resourceName}\"", innerException);