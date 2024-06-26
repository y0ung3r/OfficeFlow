﻿namespace OfficeFlow.DocumentObjectModel.Tests.Fakes;

internal sealed class FakeElement : Element
{
    public FakeElement(CompositeElement? parent = null)
        => parent?.AppendChild(this);
}