using System.IO;

namespace OfficeFlow.OpenXml.Packaging.Interfaces;

internal interface IPackageFlushStrategy
{
    void Flush(MemoryStream internalStream);
}