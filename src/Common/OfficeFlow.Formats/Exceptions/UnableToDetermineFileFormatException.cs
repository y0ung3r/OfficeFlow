using System;

namespace OfficeFlow.Formats.Exceptions;

public sealed class UnableToDetermineFileFormatException(string message) : Exception(message)
{
    public UnableToDetermineFileFormatException()
        : this("Unable to determine file format")
    { }
}