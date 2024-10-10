using Xunit;

namespace V77ApplicationWebApi.Infrastructure.Helpers.UnitTests;

public class MessageHelperTests
{
    [Fact]
    public void BuildArgsString_ShouldConcatenateArgs()
    {
        // Arrange
        string[] args = ["First", "Second", "Third", null];
        string expected = $"{{{args[0]}}} {{{args[1]}}} {{{args[2]}}} {{null}}";

        // Act
        string actual = MessageHelper.BuildArgsString(args);

        // Assert
        Assert.Equal(expected, actual);
    }
}
