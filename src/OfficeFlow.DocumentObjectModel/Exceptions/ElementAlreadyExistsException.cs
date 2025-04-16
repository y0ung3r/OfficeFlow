using System;

namespace OfficeFlow.DocumentObjectModel.Exceptions;

public sealed class ElementAlreadyExistsException(string message) 
    : Exception(message);