using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using OfficeFlow.Word.OpenXml.Packaging.Interfaces;
using OfficeFlow.Word.OpenXml.Packaging.Parts;

namespace OfficeFlow.Word.OpenXml.Packaging
{
    public sealed class OpenXmlPackage : IDisposable
    {
        public static OpenXmlPackage Create()
        {
            var flushStrategy = new PreserveFlush();
            var internalStream = new MemoryStream();
            return Load(flushStrategy, internalStream, FileMode.Create);
        }

        public static OpenXmlPackage Open(Stream remoteStream)
        {
            var flushStrategy = new FlushUsingStream(remoteStream);
            
            var internalStream = new MemoryStream();
            remoteStream.Seek(offset: 0, SeekOrigin.Begin);
            remoteStream.CopyTo(internalStream);
            
            return Load(flushStrategy, internalStream, FileMode.Open);
        }

        public static OpenXmlPackage Open(string filePath)
        {
            var flushStrategy = new FlushUsingFilePath(filePath);
            
            using var fileStream = 
                new FileStream(filePath, FileMode.Open, FileAccess.Read);
            
            var internalStream = new MemoryStream();
            fileStream.CopyTo(internalStream);
            
            return Load(flushStrategy, internalStream, FileMode.Open);
        }

        private static OpenXmlPackage Load(
            IPackageFlushStrategy flushStrategy,
            MemoryStream internalStream,
            FileMode mode)
            => new OpenXmlPackage(
                source: Package.Open(internalStream, mode, FileAccess.ReadWrite), 
                flushStrategy,
                internalStream);

        private bool _isDisposed;
        private Package _source;
        private IPackageFlushStrategy _flushStrategy;
        private readonly MemoryStream _internalStream;

        private OpenXmlPackage(
            Package source, 
            IPackageFlushStrategy flushStrategy,
            MemoryStream internalStream)
        {
            _source = source;
            _flushStrategy = flushStrategy;
            _internalStream = internalStream;
        }

        public IEnumerable<OpenXmlPackagePart> EnumerateParts()
        {
            ThrowIfDisposed();

            return _source
                .GetParts()
                .Select(OpenXmlPackagePart.Load);
        }

        public void Save()
        {
            ThrowIfDisposed();
            
            Debug.WriteIf(
                _flushStrategy is PreserveFlush, 
                $"To save a new package, you must use the {nameof(SaveTo)} method");

            foreach (var packagePart in EnumerateParts())
            {
                packagePart.Flush();
            }

            Flush();
        }

        public void SaveTo(Stream remoteStream)
        {
            ThrowIfDisposed();
            
            SetFlushStrategy(
                new FlushUsingStream(remoteStream));
            
            Save();
        }

        public void SaveTo(string filePath)
        {
            ThrowIfDisposed();
            
            SetFlushStrategy(
                new FlushUsingFilePath(filePath));
            
            Save();
        }

        public void Close() 
            => Dispose();

        /// <inheritdoc />
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }
            
            Save();
            
            _source.Close();
            _internalStream.Dispose();

            _isDisposed = true;
        }

        /// <remarks>https://github.com/dotnet/runtime/issues/24149</remarks>
        private void Flush()
        {
            _source.Close(); // Fill _internalStream & close package

            _flushStrategy.Flush(_internalStream); // Save to an external data source (remoteStream or filePath)
            
            _source = Package.Open(_internalStream, FileMode.Open, FileAccess.ReadWrite); // Reopen package
        }

        private void SetFlushStrategy(IPackageFlushStrategy flushStrategy)
            => _flushStrategy = flushStrategy;
        
        private void ThrowIfDisposed()
        {
            if (_isDisposed)
            {
                throw new ObjectDisposedException(
                    nameof(OpenXmlPackage));
            }
        }
    }
}