using System.IO;

namespace OfficeFlow.Word.OpenXml.Packaging.Interfaces
{
    internal interface IPackageFlushStrategy
    {
        void Flush(MemoryStream internalStream);
    }
}