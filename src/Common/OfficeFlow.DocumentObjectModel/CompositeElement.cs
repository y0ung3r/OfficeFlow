using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using OfficeFlow.DocumentObjectModel.Exceptions;

namespace OfficeFlow.DocumentObjectModel
{
    public abstract class CompositeElement : Element, IEnumerable<Element?>
    {
        private readonly ElementCollection _children;

        public Element? FirstChild
            => _children.Head;

        public Element? LastChild
            => _children.Tail;

        protected CompositeElement(CompositeElement? parent = null)
            : base(parent)
            => _children = new ElementCollection(this);
        
        public void AppendChild(Element element)
            => _children.Append(element);

        public void InsertAfter(Element target, Element element)
            => _children.Insert(
                GetIndexOrThrow(target) + 1, 
                element);

        public void InsertBefore(Element target, Element element)
            => _children.Insert(
                GetIndexOrThrow(target), 
                element);
        
        public void PrependChild(Element element)
            => _children.Prepend(element);

        public IEnumerable<Element> GetDescendants()
            => GetDescendants<Element>();

        public IEnumerable<TElement> GetDescendants<TElement>()
            where TElement : Element
        {
            var descedants = new Stack<Element?>(_children);
            
            while (descedants.Any())
            {
                var nextDescedant = descedants.Pop();

                if (nextDescedant is TElement element)
                {
                    yield return element;
                }

                // ReSharper disable once InvertIf
                if (nextDescedant is CompositeElement composite)
                {
                    foreach (var child in composite)
                    {
                        descedants.Push(child);
                    }
                }
            }
        }
        
        public void RemoveChild(Element element)
            => _children.Remove(element);

        public void RemoveChildren()
            => _children.Clear();

        public IEnumerator<Element?> GetEnumerator()
            => _children.GetEnumerator();

        [ExcludeFromCodeCoverage]
        IEnumerator IEnumerable.GetEnumerator()
            => GetEnumerator();

        private int GetIndexOrThrow(Element element)
        {
            var index = _children.IndexOf(element);

            if (index is -1)
            {
                throw new ElementNotFoundException(
                    $"The specified element {element.GetType().Name} is not part of the current composite element {GetType().Name}");
            }

            return index;
        }
    }
}