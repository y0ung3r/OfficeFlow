using System.Collections.Generic;
using JetBrains.Annotations;

namespace OfficeFlow.DocumentObjectModel
{
    public abstract class Element
    {
        [CanBeNull] 
        public Element PreviousSibling { get; internal set; }
        
        [CanBeNull] 
        public Element NextSibling { get; internal set; }
        
        [CanBeNull] 
        public CompositeElement Parent { get; internal set; }

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

        [CanBeNull]
        public CompositeElement GetRootElement()
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