using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using OfficeFlow.DocumentObjectModel.Exceptions;

namespace OfficeFlow.DocumentObjectModel;

public class ElementCollection : IEnumerable<Element>
{
    private int _generation;

    private readonly CompositeElement? _parent;

    public int Count { get; private set; }

    public Element? Head { get; private set; }

    public Element? Tail { get; private set; }

    public ElementCollection()
    { }

    public ElementCollection(CompositeElement parent)
        => _parent = parent;

    public Element this[int index]
    {
        get
        {
            if (index == 0 && Head is not null)
            {
                return Head;
            }

            if (index == Count - 1 && Tail is not null)
            {
                return Tail;
            }

            if (index < 0)
            {
                index = Count + index;
            }

            var skip = 0;

            foreach (var element in this)
            {
                if (skip == index)
                {
                    return element;
                }

                skip++;
            }

            throw new ElementNotFoundException(
                "The element is not present in the current collection");
        }
    }

    public bool Contains(Element element)
        => IndexOf(element) != -1;

    public int IndexOf(Element element)
    {
        if (Head == element)
        {
            return 0;
        }

        if (Tail == element)
        {
            return Count - 1;
        }

        var index = 0;
        var target = Head;

        while (target != null && target != element)
        {
            target = target.NextSibling;
            index++;
        }

        if (index < Count)
        {
            return index;
        }

        return -1;
    }

    public void Append(Element element)
        => Insert(index: Count, element);

    public void Prepend(Element element)
        => Insert(index: 0, element);

    public void Insert(int index, Element element)
    {
        if (index < 0 || index > Count)
        {
            throw new ArgumentOutOfRangeException(
                nameof(index));
        }

        if (Contains(element))
        {
            throw new ElementAlreadyExistsException(
                "The element already exists in the current collection");
        }

        element.Parent?.RemoveChild(element);
        element.Parent = _parent;

        if (Count is 0)
        {
            InsertAsFirstElement(element);

            return;
        }

        if (index == 0)
        {
            InsertBeforeFirstElement(element);
        }
        else if (index == Count)
        {
            InsertAfterLastElement(element);
        }
        else if (index < Count)
        {
            InsertAfter(this[index], element);
        }

        _generation++;
        Count++;
    }

    public void Remove(Element element)
    {
        var index = IndexOf(element);

        if (index == -1)
        {
            return;
        }

        RemoveAt(index);
    }

    public void RemoveAt(int index)
    {
        var element = this[index];

        if (element.PreviousSibling != null)
        {
            element.PreviousSibling.NextSibling = element.NextSibling;
        }

        if (element.NextSibling != null)
        {
            element.NextSibling.PreviousSibling = element.PreviousSibling;
        }

        if (Head == element)
        {
            Head = element.NextSibling;
        }

        if (Tail == element)
        {
            Tail = element.PreviousSibling;
        }

        element.PreviousSibling = null;
        element.NextSibling = null;
        element.Parent = null;

        _generation++;
        Count--;
    }

    public void Clear()
    {
        var root = Head;

        while (root != null)
        {
            var next = root.NextSibling;
            
            Remove(root);
            
            root = next;
        }

        Head = null;
        Tail = null;

        _generation++;
        Count = 0;
    }

    public IEnumerator<Element> GetEnumerator()
        => new ElementIterator(this);

    [ExcludeFromCodeCoverage]
    IEnumerator IEnumerable.GetEnumerator()
        => GetEnumerator();

    private void InsertAsFirstElement(Element element)
    {
        element.NextSibling = null;
        element.PreviousSibling = null;

        Head = element;
        Tail = element;

        _generation++;
        Count++;
    }

    private void InsertBeforeFirstElement(Element element)
    {
        element.NextSibling = Head;

        if (Head != null)
        {
            Head.PreviousSibling = element;
        }

        Head = element;
    }

    private void InsertAfterLastElement(Element element)
    {
        element.PreviousSibling = Tail;

        if (Tail != null)
        {
            Tail.NextSibling = element;
        }

        Tail = element;
    }

    private void InsertAfter(Element existsElement, Element element)
    {
        element.NextSibling = existsElement;
        element.PreviousSibling = existsElement.PreviousSibling;

        if (existsElement.PreviousSibling != null)
        {
            existsElement.PreviousSibling.NextSibling = element;
        }

        if (existsElement.NextSibling != null)
        {
            existsElement.NextSibling.PreviousSibling = existsElement;
        }

        existsElement.PreviousSibling = element;
    }

    private sealed class ElementIterator : IEnumerator<Element>
    {
        private bool _isDisposed;
        
        private readonly ElementCollection _elements;

        private readonly int _generation;

        private Element? _next;

        private Element? _current;

        private int _index;

        public Element Current
        {
            get
            {
                if (_index is 0 || _index == _elements.Count + 1)
                {
                    throw new InvalidOperationException(
                        "Index was outside the bounds of the collection");
                }

                return _current
                    ?? throw new InvalidOperationException(
                        "Do not call Current before invoking MoveNext");
            }
        }

        [ExcludeFromCodeCoverage]
        object IEnumerator.Current
            => Current;

        public ElementIterator(ElementCollection elements)
        {
            _elements = elements;
            _generation = _elements._generation;
            _next = _elements.Head;
            _current = null;
            _index = 0;
        }

        public bool MoveNext()
        {
            ThrowIfDisposed();
            ThrowIfCollectionWasModified();

            if (_next is null)
            {
                _index = _elements.Count + 1;

                return false;
            }

            _current = _next;
            _next = _next.NextSibling;

            if (_next == _elements.Head)
            {
                _next = null;
            }

            _index++;

            return true;
        }

        public void Reset()
        {
            ThrowIfDisposed();
            ThrowIfCollectionWasModified();

            _current = null;
            _next = _elements.Head;
            _index = 0;
        }

        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _isDisposed = true;
        }

        private void ThrowIfDisposed()
        {
            if (!_isDisposed)
                return;

            throw new ObjectDisposedException(
                nameof(ElementIterator));
        }
        
        private void ThrowIfCollectionWasModified()
        {
            if (_generation != _elements._generation)
            {
                throw new InvalidOperationException(
                    "The collection has been modified");
            }
        }
    }
}