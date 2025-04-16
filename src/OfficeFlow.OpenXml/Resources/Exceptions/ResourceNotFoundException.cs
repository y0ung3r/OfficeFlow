using System;

namespace OfficeFlow.OpenXml.Resources.Exceptions;

public sealed class ResourceNotFoundException(string resourceName, Exception? innerException = null)
    : Exception($"Resource \"{resourceName}\" not found", innerException);