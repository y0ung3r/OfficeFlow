namespace OfficeFlow.DocumentObjectModel.Tests.Fakes;

internal sealed class FakeCustomElement : CompositeElement 
{
    public FakeCustomElement(CompositeElement? parent = null)
        => parent?.AppendChild(this);
}