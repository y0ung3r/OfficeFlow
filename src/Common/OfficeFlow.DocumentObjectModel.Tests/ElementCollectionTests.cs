using System;
using FluentAssertions;
using OfficeFlow.DocumentObjectModel.Exceptions;
using OfficeFlow.DocumentObjectModel.Tests.Fakes;
using Xunit;

namespace OfficeFlow.DocumentObjectModel.Tests;

public sealed class ElementCollectionTests : DocumentObjectModelTestsBase
{
    private readonly ElementCollection _sut = new();
    private readonly FakeElement _head = new();
    private readonly FakeElement _second = new();
    private readonly FakeElement _third = new();
    private readonly FakeElement _fourth = new();
    private readonly FakeElement _tail = new();
    private readonly FakeElement _detached = new();

    [Fact]
    public void Should_returns_index_of_element()
    {
        // Arrange
        _sut.Append(_head);
        _sut.Append(_second);
        _sut.Append(_tail);

        // Act & Assert
        _sut.IndexOf(_detached)
            .Should()
            .Be(-1);

        _sut.IndexOf(_head)
            .Should()
            .Be(0);

        _sut.IndexOf(_second)
            .Should()
            .Be(1);

        _sut.IndexOf(_tail)
            .Should()
            .Be(2);
    }

    [Fact]
    public void Should_returns_element_at_index_from_start()
    {
        // Arrange
        _sut.Append(_head);
        _sut.Append(_second);
        _sut.Append(_tail);

        // Act
        var target = _sut[index: 1];

        // Assert
        VerifyExistenceInOrder(_head, target, _tail);
    }

    [Fact]
    public void Should_returns_element_at_index_from_end()
    {
        // Arrange
        _sut.Append(_head);
        _sut.Append(_second);
        _sut.Append(_tail);

        // Act
        var target = _sut[index: -2];

        // Assert
        VerifyExistenceInOrder(_head, target, _tail);
    }

    [Fact]
    public void Should_returns_null_if_element_is_not_presented_in_collection()
        => _sut[-5]
            .Should()
            .BeNull();

    [Fact]
    public void Should_append_element_properly()
    {
        // Act
        _sut.Append(_head);
        _sut.Append(_second);
        _sut.Append(_tail);

        // Assert
        _sut.Count
            .Should()
            .Be(3);

        VerifyExistenceInOrder(_head, _second, _tail);
    }

    [Fact]
    public void Should_insert_element_properly()
    {
        // Arrange
        _sut.Append(_head);
        _sut.Append(_second);
        _sut.Append(_tail);

        // Act
        _sut.Insert(index: 2, _third);

        // Assert
        _sut.Count
            .Should()
            .Be(4);

        VerifyExistenceInOrder(_head, _second, _third, _tail);
    }

    [Fact]
    public void Should_prepend_element_properly()
    {
        // Act
        _sut.Prepend(_tail);
        _sut.Prepend(_second);
        _sut.Prepend(_head);

        // Assert
        _sut.Count
            .Should()
            .Be(3);

        VerifyExistenceInOrder(_head, _second, _tail);
    }

    [Fact]
    public void Should_remove_element_properly()
    {
        // Arrange
        _sut.Append(_head);
        _sut.Append(_second);
        _sut.Append(_tail);

        // Act
        _sut.Remove(_second);

        // Assert
        _sut.Count
            .Should()
            .Be(2);

        VerifyRemoving(_second);
        VerifyExistenceInOrder(_head, _tail);
    }
        
    [Fact]
    public void Should_do_nothing_if_removable_element_is_not_presented_in_collection()
    {
        // Arrange
        _sut.Append(_head);
        _sut.Append(_second);
        _sut.Append(_tail);

        // Act
        _sut.Remove(_third);
            
        // Assert
        _sut.Count
            .Should()
            .Be(3);
            
        VerifyExistenceInOrder(_head, _second, _tail);
    }

    [Fact]
    public void Should_remove_element_at_index_properly()
    {
        // Arrange
        _sut.Append(_head);
        _sut.Append(_second);
        _sut.Append(_tail);

        // Act
        _sut.RemoveAt(index: 1);

        // Assert
        _sut.Count
            .Should()
            .Be(2);

        VerifyRemoving(_second);
        VerifyExistenceInOrder(_head, _tail);
    }

    [Theory]
    [InlineData(-5)]
    [InlineData(-15)]
    [InlineData(5)]
    [InlineData(15)]
    public void Should_throws_exception_while_removing_nonexistent_element(int index)
    {
        // Arrange
        _sut.Append(_head);
        
        // Act
        var action = () => _sut.RemoveAt(index);
            
        // Assert
        action
            .Should()
            .Throw<ElementNotFoundException>();
    }

    [Fact]
    public void Should_clear_collection()
    {
        // Arrange
        var elements =
            new[]
            {
                _head,
                _second,
                _third,
                _fourth,
                _tail
            };

        foreach (var element in elements)
        {
            _sut.Append(element);
        }

        // Act
        _sut.Clear();

        // Assert
        elements
            .Should()
            .AllSatisfy(VerifyRemoving);

        _sut.Count
            .Should()
            .Be(0);
    }

    [Fact]
    public void Should_iterate_elements_in_strict_order()
    {
        // Arrange
        var elements =
            new[]
            {
                _head,
                _second,
                _third,
                _fourth,
                _tail
            };

        foreach (var element in elements)
        {
            _sut.Append(element);
        }

        // Act & Assert
        _sut.Should()
            .BeEquivalentTo(elements,
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void Should_throws_exception_if_attempts_append_the_exists_element()
    {
        // Arrange
        _sut.Append(_head);
        
        // Act
        var action = () => _sut.Append(_head);

        // Assert
        action
            .Should()
            .Throw<ElementAlreadyExistsException>();
    }

    [Fact]
    public void Should_throws_exception_if_attempts_prepend_the_exists_element()
    {
        // Arrange
        _sut.Append(_head);
        
        // Act
        var action = () => _sut.Prepend(_head);

        // AssertAssert
        action
            .Should()
            .Throw<ElementAlreadyExistsException>();
    }

    [Theory]
    [InlineData(-5)]
    [InlineData(-15)]
    [InlineData(5)]
    [InlineData(15)]
    public void Should_throws_exception_if_attempts_insert_element_to_out_of_the_range(int index)
    {
        // Arrange
        _sut.Append(_head);
        
        // Act
        var action = () => _sut.Insert(index, _second);

        // Assert
        action
            .Should()
            .Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Should_throws_exception_if_attempts_iterate_the_modified_collection()
    {
        // Arrange
        var elements = new[]
        {
            _head,
            _second
        };

        foreach (var element in elements)
        {
            _sut.Append(element);
        }

        // Act
        var action = () =>
        {
            foreach (var _ in _sut)
            {
                _sut.Append(_tail);
            }
        };

        // Assert
        action
            .Should()
            .Throw<InvalidOperationException>();
    }
}