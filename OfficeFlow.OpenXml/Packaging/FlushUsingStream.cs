using System.IO;
using OfficeFlow.OpenXml.Packaging.Interfaces;

namespace OfficeFlow.OpenXml.Packaging
{
    internal sealed class FlushUsingStream : IPackageFlushStrategy
    {
        private readonly Stream _remoteStream;

        public FlushUsingStream(Stream remoteStream)
            => _remoteStream = remoteStream;
        
        /// <inheritdoc />
        public void Flush(MemoryStream internalStream)
        {
            if (_remoteStream == internalStream)
            {
                return;
            }
            
            _remoteStream.SetLength(value: 0);
            _remoteStream.Seek(offset: 0, SeekOrigin.Begin);
            
            internalStream.WriteTo(_remoteStream);
        }
    }
}