﻿namespace OfficeFlow.DocumentObjectModel.Tests.Fakes;

internal sealed class FakeCompositeElement : CompositeElement
{
    public FakeCompositeElement(CompositeElement? parent = null)
        => parent?.AppendChild(this);
}