﻿using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeFlow.TestFramework;

/// <summary>
/// Clearing temporary files when tests are failed or passed
/// </summary>
public sealed class TempFilePool : IDisposable
{
    private readonly List<string> _filePaths = [];

    public string GetTempFilePath()
    {
        var filePath = Path.GetTempFileName();

        _filePaths.Add(filePath);

        return filePath;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        foreach (var filePath in _filePaths)
        {
            File.Delete(filePath);
        }
    }
}