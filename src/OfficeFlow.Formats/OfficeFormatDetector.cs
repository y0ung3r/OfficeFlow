using System.IO;
using System.Linq;
using OfficeFlow.Formats.Exceptions;
using OfficeFlow.Formats.Interfaces;

namespace OfficeFlow.Formats;

public static class OfficeFormatDetector
{
    private static readonly IOfficeFormat[] Formats =
    {
        new OpenXmlFormat(),
        new BinaryFormat()
    };

    private static readonly int MaxHexdumpLength
        = Formats.Max(signature =>
            signature
                .Hexdumps
                .Max(hexdump => hexdump.Length));

    public static IOfficeFormat Detect(Stream stream)
    {
        ThrowIfUnableToDetermineFileFormat(stream);

        using var binaryReader = 
            new BinaryReader(stream);
        
        var hexdump = binaryReader.ReadBytes(MaxHexdumpLength);

        return Formats.FirstOrDefault(format =>
            format.Hexdumps.Any(hexdump.SequenceEqual))
               ?? throw new UnableToDetermineFileFormatException();
    }

    public static IOfficeFormat Detect(string filePath)
    {
        ThrowIfUnableToDetermineFileFormat(filePath);

        var extension = Path.GetExtension(filePath);
        var officeFormat = Formats.FirstOrDefault(format =>
            format.Extensions.Contains(extension));

        if (officeFormat != null)
            return officeFormat;

        using var stream = 
            File.OpenRead(filePath);
        
        return Detect(stream);
    }

    private static void ThrowIfUnableToDetermineFileFormat(Stream? stream)
    {
        if (stream is null)
        {
            throw new UnableToDetermineFileFormatException(
                $"Unable to determine file format, because the specified {nameof(Stream)} is not initialized");
        }

        if (stream.CanRead is false)
        {
            throw new UnableToDetermineFileFormatException(
                "Unable to determine file format, because the file cannot be read");
        }

        if (stream.Length == 0)
        {
            throw new UnableToDetermineFileFormatException(
                "Unable to determine file format, since the specified file is empty");
        }
    }

    private static void ThrowIfUnableToDetermineFileFormat(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new UnableToDetermineFileFormatException(
                $"Unable to determine file format from the specified path: \"{filePath}\"");
        }
    }
}