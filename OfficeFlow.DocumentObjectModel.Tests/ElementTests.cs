using System.Linq;
using FluentAssertions;
using OfficeFlow.DocumentObjectModel.Tests.Fakes;
using Xunit;

namespace OfficeFlow.DocumentObjectModel.Tests
{
    public sealed class ElementTests
    {
        [Fact]
        public void Should_detach_element_from_old_parent()
        {
            // Arrange
            var firstParent = new FakeCompositeElement();
            var secondParent = new FakeCompositeElement();
            var sut = new FakeElement(firstParent);
            
            // Act
            secondParent.AppendChild(sut);

            // Assert
            firstParent
                .Should()
                .NotContain(sut);

            secondParent
                .Should()
                .Contain(sut);
        }
        
        [Fact]
        public void Should_add_element_to_parent_from_constructor()
        {
            // Arrange
            var parent = new FakeCompositeElement();
            
            // Act
            var sut = new FakeElement(parent);

            // Assert
            parent
                .Should()
                .Contain(sut);
        }
        
        [Fact]
        public void Should_returns_root_element_if_exists()
        {
            // Arrange
            var root = new FakeCompositeElement();
            var rootChild = new FakeCompositeElement(root);
            var parent = new FakeCompositeElement(rootChild);
            var sut = new FakeElement(parent);

            // Act
            var actualRoot = sut.GetRootElement();

            // Assert
            actualRoot
                .Should()
                .BeSameAs(root);
        }

        [Fact]
        public void Should_returns_null_root_element_if_not_exists()
        {
            // Arrange
            var sut = new FakeElement();

            // Act
            var root = sut.GetRootElement();

            // Assert
            root
                .Should()
                .BeNull();
        }

        [Fact]
        public void Should_returns_all_ancestors()
        {
            // Arrange
            var root = new FakeCompositeElement();
            var rootChild = new FakeCompositeElement(root);
            var parent = new FakeCompositeElement(rootChild);
            var sut = new FakeElement(parent);
            
            // Act
            var ancestors = sut.GetAncestors();
            
            // Assert
            ancestors
                .Should()
                .ContainInOrder(
                    parent,
                    rootChild,
                    root);
        }

        [Fact]
        public void Should_returns_ancestors_of_type()
        {
            // Arrange
            var root = new FakeCompositeElement();
            var firstRootChild = new FakeCompositeElement(root);
            var secondRootChild = new FakeCustomElement(firstRootChild);
            var parent = new FakeCustomElement(secondRootChild);
            var sut = new FakeElement(parent);
            
            // Act
            var ancestors = sut
                .GetAncestors<FakeCustomElement>()
                .ToArray();
            
            // Assert
            ancestors
                .Should()
                .ContainInOrder(
                    parent,
                    secondRootChild);

            ancestors
                .Should()
                .AllBeOfType<FakeCustomElement>();
        }
    }
}