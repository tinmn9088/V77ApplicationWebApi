using V77ApplicationWebApi.Infrastructure.Helpers;
using Xunit;

namespace V77ApplicationWebApi.Infrastructure.UnitTests.Helpers;

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
