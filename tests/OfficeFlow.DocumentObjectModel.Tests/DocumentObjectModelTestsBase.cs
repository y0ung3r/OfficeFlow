using FluentAssertions;

namespace OfficeFlow.DocumentObjectModel.Tests;

public abstract class DocumentObjectModelTestsBase
{
    protected void VerifyRemoving(Element element)
    {
        element
            .NextSibling
            .Should()
            .BeNull();

        element
            .PreviousSibling
            .Should()
            .BeNull();
    }

    protected void VerifyExistenceInOrder(params Element?[] elements)
    {
        for (var index = 0; index < elements.Length - 1; index++)
        {
            var element = elements[index];

            if (index is 0)
            {
                element?
                    .PreviousSibling
                    .Should()
                    .BeNull();

                element?
                    .NextSibling
                    .Should()
                    .Be(elements[1]);

                continue;
            }

            if (index == elements.Length - 1)
            {
                element?
                    .PreviousSibling
                    .Should()
                    .Be(elements[^2]);

                element?
                    .NextSibling
                    .Should()
                    .BeNull();

                continue;
            }

            element?
                .PreviousSibling
                .Should()
                .Be(elements[index - 1]);

            element?
                .NextSibling
                .Should()
                .Be(elements[index + 1]);
        }
    }
}