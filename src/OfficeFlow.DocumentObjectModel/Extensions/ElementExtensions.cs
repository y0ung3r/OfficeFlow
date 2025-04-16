using OfficeFlow.DocumentObjectModel.Exceptions;

namespace OfficeFlow.DocumentObjectModel.Extensions;

public static class ElementExtensions
{
    public static void InsertAfterSelf(this Element self, Element target)
        => self
            .GetParentOrThrow()
            .InsertAfter(self, target);

    public static void InsertBeforeSelf(this Element self, Element target)
        => self
            .GetParentOrThrow()
            .InsertBefore(self, target);

    private static CompositeElement GetParentOrThrow(this Element self)
        => self.Parent ?? throw new ElementHasNoParentException();
}