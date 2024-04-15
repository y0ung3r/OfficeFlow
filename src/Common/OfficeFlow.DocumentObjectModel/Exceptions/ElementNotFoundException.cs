using System;

namespace OfficeFlow.DocumentObjectModel.Exceptions;

public sealed class ElementNotFoundException(string message) 
    : Exception(message);