using System.Xml.Linq;
using FluentAssertions;
using FluentAssertions.Xml;

namespace OfficeFlow.TestFramework.Extensions;

public static class XElementAssertionsExtensions
{
    public static AndConstraint<XElementAssertions> HaveName(this XElementAssertions assertions, XName name)
        => assertions.Match(element => element.Name == name, because: "XElement must have a specified name: \"{0}\"", name);
    
    public static AndConstraint<XElementAssertions> NotHaveName(this XElementAssertions assertions, XName name)
        => assertions.Match(element => element.Name != name, because: "XElement must not have a specified name: \"{0}\"", name);
}