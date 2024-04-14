using System.Linq;
using FluentAssertions;
using OfficeFlow.DocumentObjectModel.Exceptions;
using OfficeFlow.DocumentObjectModel.Tests.Fakes;
using Xunit;

namespace OfficeFlow.DocumentObjectModel.Tests;

public sealed class CompositeElementTests : DocumentObjectModelTestsBase
{
    private readonly FakeElement _head = new();
    private readonly FakeElement _second = new();
    private readonly FakeElement _third = new();
    private readonly FakeElement _tail = new();

    [Fact]
    public void Should_returns_first_and_last_child()
    {
        // Arrange
        var sut = new FakeCompositeElement();
            
        // Act
        sut.AppendChild(_head);
        sut.AppendChild(_tail);

        // Assert
        sut.FirstChild
            .Should()
            .Be(_head);
            
        sut.LastChild
            .Should()
            .Be(_tail);
    }

    [Fact]
    public void Should_append_children()
    {
        // Arrange
        var sut = new FakeCompositeElement();

        // Act
        sut.AppendChild(_head);
        sut.AppendChild(_second);
        sut.AppendChild(_tail);

        // Assert
        VerifyExistenceInOrder(_head, _second, _tail);
    }

    [Fact]
    public void Should_prepend_children()
    {
        // Arrange
        var sut = new FakeCompositeElement();

        // Act
        sut.PrependChild(_tail);
        sut.PrependChild(_second);
        sut.PrependChild(_head);

        // Assert
        VerifyExistenceInOrder(_head, _second, _tail);
    }

    [Fact]
    public void Should_insert_child_after_existing_element()
    {
        // Arrange
        var sut = new FakeCompositeElement();

        sut.AppendChild(_head);
        sut.AppendChild(_second);
        sut.AppendChild(_tail);

        // Act
        sut.InsertAfter(_second, _third);

        // Assert
        VerifyExistenceInOrder(_head, _second, _third, _tail);
    }

    [Fact]
    public void Should_throws_exception_while_insert_child_after_existing_element_if_not_exists()
        => new FakeCompositeElement()
            .Invoking(sut => sut.InsertAfter(_head, _second))
            .Should()
            .Throw<ElementNotFoundException>();

    [Fact]
    public void Should_insert_child_before_existing_element()
    {
        // Arrange
        var sut = new FakeCompositeElement();

        sut.AppendChild(_head);
        sut.AppendChild(_third);
        sut.AppendChild(_tail);

        // Act
        sut.InsertBefore(_third, _second);

        // Assert
        VerifyExistenceInOrder(_head, _second, _third, _tail);
    }
        
    [Fact]
    public void Should_throws_exception_while_insert_child_before_existing_element_if_not_exists()
        => new FakeCompositeElement()
            .Invoking(sut => sut.InsertBefore(_head, _second))
            .Should()
            .Throw<ElementNotFoundException>();

    [Fact]
    public void Should_returns_all_descendants()
    {
        // Arrange
        var sut = new FakeCompositeElement();
        var firstChild = new FakeCompositeElement(sut);
        var secondChild = new FakeCompositeElement(firstChild);
        var lastChild = new FakeElement(secondChild);
            
        // Act
        var descendants = sut.GetDescendants();
            
        // Assert
        descendants
            .Should()
            .ContainInOrder(
                firstChild,
                secondChild,
                lastChild);
    }

    [Fact]
    public void Should_returns_descendants_of_type()
    {
        // Arrange
        var sut = new FakeCompositeElement();
        var firstChild = new FakeCustomElement(sut);
        var secondChild = new FakeCompositeElement(firstChild);
        var lastChild = new FakeCustomElement(secondChild);
            
        // Act
        var ancestors = sut
            .GetDescendants<FakeCustomElement>()
            .ToArray();
            
        // Assert
        ancestors
            .Should()
            .ContainInOrder(
                firstChild,
                lastChild);

        ancestors
            .Should()
            .AllBeOfType<FakeCustomElement>();
    }

    [Fact]
    public void Should_remove_children()
    {
        // Arrange
        var sut = new FakeCompositeElement();
        var elements = new[]
        {
            _head,
            _second,
            _tail
        };

        foreach (var element in elements)
        {
            sut.AppendChild(element);
        }

        // Act
        sut.RemoveChildren();

        // Assert
        elements
            .Should()
            .AllSatisfy(VerifyRemoving);

        sut.Should()
            .BeEmpty();
    }
}