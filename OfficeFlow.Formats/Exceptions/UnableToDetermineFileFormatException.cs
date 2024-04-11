using System;

namespace OfficeFlow.Formats.Exceptions
{
    public sealed class UnableToDetermineFileFormatException : Exception
    {
        public UnableToDetermineFileFormatException()
            : this("Unable to determine file format")
        { }

        public UnableToDetermineFileFormatException(string message)
            : base(message)
        { }
    }
}