using System.Collections.Generic;

namespace OfficeFlow.DocumentObjectModel
{
    public abstract class Element
    {
        public Element? PreviousSibling { get; internal set; }
        
        public Element? NextSibling { get; internal set; }
        
        public CompositeElement? Parent { get; internal set; }

        public IEnumerable<Element> GetAncestors()
            => GetAncestors<Element>();

        public IEnumerable<TElement> GetAncestors<TElement>()
            where TElement : Element
        {
            var nextAncestor = Parent;

            while (nextAncestor != null)
            {
                if (nextAncestor is TElement element)
                {
                    yield return element;
                }

                nextAncestor = nextAncestor.Parent;
            }
        }

        public CompositeElement? GetRootElement()
        {
            var root = Parent;

            while (root?.Parent != null)
            {
                root = root.Parent;
            }

            return root;
        }
    }
}