using System;
using FluentAssertions;
using OfficeFlow.DocumentObjectModel.Exceptions;
using OfficeFlow.DocumentObjectModel.Extensions;
using OfficeFlow.DocumentObjectModel.Tests.Fakes;
using Xunit;

namespace OfficeFlow.DocumentObjectModel.Tests
{
    public sealed class ElementExtensionsTests : DocumentObjectModelTestsBase
    {
        private readonly FakeElement _head
            = new FakeElement();
        
        private readonly FakeElement _second
            = new FakeElement();
        
        private readonly FakeElement _third
            = new FakeElement();
        
        private readonly FakeElement _tail
            = new FakeElement();
        
        [Fact]
        public void Should_insert_element_after_self()
        {
            // Arrange
            var parent = new FakeCompositeElement();

            parent.AppendChild(_head);
            parent.AppendChild(_second);
            parent.AppendChild(_tail);

            // Act
            _second.InsertAfterSelf(_third);

            // Assert
            VerifyExistenceInOrder(_head, _second, _third, _tail);
        }
        
        [Fact]
        public void Should_throws_exception_while_insert_element_after_self_if_has_no_parent()
        {
            // Arrange & Act & Assert
            new Action(() => _third.InsertAfterSelf(_third))
                .Should()
                .Throw<ElementHasNoParentException>();
        }

        [Fact]
        public void Should_insert_element_before_self()
        {
            // Arrange
            var parent = new FakeCompositeElement();

            parent.AppendChild(_head);
            parent.AppendChild(_third);
            parent.AppendChild(_tail);

            // Act
            _third.InsertBeforeSelf(_second);

            // Assert
            VerifyExistenceInOrder(_head, _second, _third, _tail);
        }

        [Fact]
        public void Should_throws_exception_while_insert_element_before_self_if_has_no_parent()
        {
            // Arrange & Act & Assert
            new Action(() => _third.InsertBeforeSelf(_second))
                .Should()
                .Throw<ElementHasNoParentException>();
        }
    }
}