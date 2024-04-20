using System.IO;
using System.Reflection;

namespace OfficeFlow.OpenXml.Tests.Resources;

public sealed class FakeAssembly(params string[] resourceNames) : Assembly
{
    /// <inheritdoc />
    public override string[] GetManifestResourceNames()
        => resourceNames;

    /// <inheritdoc />
    public override Stream? GetManifestResourceStream(string name)
        => null;
}