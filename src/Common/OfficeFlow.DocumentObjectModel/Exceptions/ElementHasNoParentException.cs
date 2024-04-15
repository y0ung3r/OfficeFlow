using System;

namespace OfficeFlow.DocumentObjectModel.Exceptions;

public sealed class ElementHasNoParentException() 
    : Exception("Element does not have a parent");