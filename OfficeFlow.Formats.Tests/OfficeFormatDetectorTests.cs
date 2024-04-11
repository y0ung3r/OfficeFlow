using System;
using System.IO;
using FluentAssertions;
using OfficeFlow.Formats.Exceptions;
using OfficeFlow.TestFramework;
using Xunit;

namespace OfficeFlow.Formats.Tests
{
    public sealed class OfficeFormatDetectorTests : IClassFixture<TempFilePool>
    {
        private readonly TempFilePool _tempFilePool;

        public OfficeFormatDetectorTests(TempFilePool tempFilePool)
            => _tempFilePool = tempFilePool;

        [Theory]
        [InlineData("document.doc")]
        [InlineData("document.dot")]
        [InlineData("document.xls")]
        [InlineData("document.xlt")]
        [InlineData("document.xlm")]
        [InlineData("document.pps")]
        [InlineData("document.ppt")]
        [InlineData("document.pot")]
        public void Should_detect_binary_format_from_file_path(string filePath)
        {
            // Arrange & Act
            var fileFormat = OfficeFormatDetector.Detect(filePath);

            // Assert
            fileFormat
                .Should()
                .BeOfType<BinaryFormat>();
        }

        [Fact]
        public void Should_detect_binary_format_from_stream_if_file_path_contains_undefined_extension()
        {
            // Arrange
            var filePath = _tempFilePool.GetTempFilePath();
            var hexdump = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            File.WriteAllBytes(filePath, hexdump);

            // Act
            var fileFormat = OfficeFormatDetector.Detect(filePath);

            // Assert
            fileFormat
                .Should()
                .BeOfType<BinaryFormat>();
        }

        [Theory]
        [InlineData("document.docx")]
        [InlineData("document.docm")]
        [InlineData("document.dotx")]
        [InlineData("document.dotm")]
        [InlineData("document.xlsx")]
        [InlineData("document.xlsm")]
        [InlineData("document.xltx")]
        [InlineData("document.xltm")]
        [InlineData("document.pptx")]
        [InlineData("document.pptm")]
        [InlineData("document.potx")]
        [InlineData("document.potm")]
        [InlineData("document.ppsx")]
        [InlineData("document.ppsm")]
        public void Should_detect_open_xml_format_from_file_path(string filePath)
        {
            // Arrange & Act
            var fileFormat = OfficeFormatDetector.Detect(filePath);

            // Assert
            fileFormat
                .Should()
                .BeOfType<OpenXmlFormat>();
        }

        [Theory]
        [InlineData(new byte[] { 0x50, 0x4B, 0x03, 0x04 })]
        [InlineData(new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00 })]
        public void Should_detect_open_xml_format_from_stream_if_file_path_contains_undefined_extension(byte[] hexdump)
        {
            // Arrange
            var filePath = _tempFilePool.GetTempFilePath();
            File.WriteAllBytes(filePath, hexdump);

            // Act
            var fileFormat = OfficeFormatDetector.Detect(filePath);

            // Assert
            fileFormat
                .Should()
                .BeOfType<OpenXmlFormat>();
        }

        [Fact]
        public void Should_detect_binary_format_from_stream()
        {
            // Arrange
            var hexdump = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            
            using var stream = 
                new MemoryStream(hexdump);

            // Act
            var fileFormat = OfficeFormatDetector.Detect(stream);

            // Assert
            fileFormat
                .Should()
                .BeOfType<BinaryFormat>();
        }

        [Theory]
        [InlineData(new byte[] { 0x50, 0x4B, 0x03, 0x04 })]
        [InlineData(new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00 })]
        public void Should_detect_open_xml_format_from_stream(byte[] hexdump)
        {
            // Arrange
            using var stream = 
                new MemoryStream(hexdump);

            // Act
            var fileFormat = OfficeFormatDetector.Detect(stream);

            // Assert
            fileFormat
                .Should()
                .BeOfType<OpenXmlFormat>();
        }

        [Fact]
        public void Should_throw_exception_if_format_is_undefined()
            => new Action(() =>
            {
                using var stream = 
                    new MemoryStream(new byte[] { 0xD0 });
                
                OfficeFormatDetector.Detect(stream);
            })
            .Should()
            .Throw<UnableToDetermineFileFormatException>();
        
        [Fact]
        public void Should_throw_exception_if_stream_is_null()
            => new Action(() => OfficeFormatDetector.Detect(default(Stream)!))
                .Should()
                .Throw<UnableToDetermineFileFormatException>();

        [Fact]
        public void Should_throw_exception_if_stream_is_unreadable()
            => new Action(() =>
            {
                using var stream = 
                    new MemoryStream(new byte[] { 0xD0 });
                
                stream.Close();
                
                OfficeFormatDetector.Detect(stream);
            })
            .Should()
            .Throw<UnableToDetermineFileFormatException>();

        [Fact]
        public void Should_throw_exception_if_stream_is_empty()
            => new Action(() =>
            {
                using var stream = 
                    new MemoryStream(Array.Empty<byte>());
                    
                OfficeFormatDetector.Detect(stream);
            })
            .Should()
            .Throw<UnableToDetermineFileFormatException>();

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void Should_throw_exception_if_file_path_is_null_or_empty(string? filePath)
            => new Action(() => OfficeFormatDetector.Detect(filePath!))
                .Should()
                .Throw<UnableToDetermineFileFormatException>();
    }
}