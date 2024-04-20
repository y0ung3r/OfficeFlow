using System;
using FluentAssertions;
using OfficeFlow.DocumentObjectModel.Tests.Fakes;
using Xunit;

namespace OfficeFlow.DocumentObjectModel.Tests;

public sealed class ElementIteratorTests
{
    [Fact]
    public void Should_throw_exception_if_collection_was_modified()
    {
        // Arrange
        var collection = new ElementCollection();
        
        // Act & Assert
        var testCases = new Action[]
        {
            () =>
            {
                using var sut = 
                    collection.GetEnumerator();
                
                collection.Append(
                    new FakeElement());
                
                sut.MoveNext();
            },
            () =>
            {
                using var sut = 
                    collection.GetEnumerator();
                
                collection.Append(
                    new FakeElement());
                
                sut.Reset();
            }
        };

        testCases
            .Should()
            .AllSatisfy(testCase =>
                testCase
                    .Should()
                    .Throw<InvalidOperationException>());
    }
    
    [Fact]
    public void Should_throw_exception_if_iterator_disposed()
    {
        // Arrange
        var sut = new ElementCollection()
            .GetEnumerator();
        
        // Act
        sut.Dispose();
        
        // Assert
        var testCases = new Action[]
        {
            () => _ = sut.Current,
            () => sut.MoveNext(),
            () => sut.Reset()
        };

        testCases
            .Should()
            .AllSatisfy(testCase =>
                testCase
                    .Should()
                    .Throw<ObjectDisposedException>());
    }
    
    [Fact]
    public void Should_throw_exception_if_index_of_current_element_was_outside_the_bounds_of_the_collection()
    {
        // Arrange
        using var sut = new ElementCollection()
            .GetEnumerator();
        
        // Act
        sut.MoveNext();
        
        // Assert
        sut.Invoking(iterator => iterator.Current)
            .Should()
            .Throw<InvalidOperationException>();
    }

    [Fact]
    public void Should_stop_iterating_if_collection_is_empty()
    {
        // Arrange
        using var sut = new ElementCollection()
            .GetEnumerator();
        
        // Act & Assert
        sut.MoveNext()
            .Should()
            .BeFalse();
    }
    
    [Fact]
    public void Should_stop_iterating_if_the_end_is_reached()
    {
        // Arrange
        var collection = new ElementCollection();
        
        collection.Append(
            new FakeElement());
        
        using var sut = 
            collection.GetEnumerator();
        
        // Act & Assert
        sut.MoveNext()
            .Should()
            .BeTrue();
        
        sut.MoveNext()
            .Should()
            .BeFalse();
    }
    
    [Fact]
    public void Should_reset_iterator_properly()
    {
        // Arrange
        var collection = new ElementCollection();
        
        collection.Append(
            new FakeElement());
        
        using var sut = 
            collection.GetEnumerator();

        sut.MoveNext();
        sut.MoveNext();
        
        // Act
        sut.Reset();
        
        // Assert
        sut.MoveNext()
            .Should()
            .BeTrue();
    }

    [Fact]
    public void Should_do_nothing_if_dispose_is_called_more_than_once()
    {
        // Arrange
        var sut = new ElementCollection()
            .GetEnumerator();
        
        // Act
        sut.Dispose();
        
        // Assert
        sut.Invoking(iterator => iterator.Dispose())
            .Should()
            .NotThrow();
    }
}