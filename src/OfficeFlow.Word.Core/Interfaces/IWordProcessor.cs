using System;
using System.IO;
using OfficeFlow.Word.Core.Elements;

namespace OfficeFlow.Word.Core.Interfaces;

public interface IWordProcessor : IDisposable
{
    Body Body { get; }

    void Save();

    void SaveTo(Stream stream);

    void SaveTo(string filePath);
}