using System;
using Xunit;

namespace V77ApplicationWebApi.Core.UnitTests;

public class ConnectionPropertiesTests
{
    private const string TestInfobasePath = @"D:\TestInfobase";

    private const string TestUsername = "TestUsername";

    private const string TestPassword = "TestPassword";

    [Fact]
    public void Equals_ShouldComparePropertiesValues()
    {
        // Arrange
        ConnectionProperties properties1 = new(
            infobasePath: TestInfobasePath,
            username: TestUsername,
            password: TestPassword);
        ConnectionProperties properties2 = new(
            infobasePath: TestInfobasePath,
            username: TestUsername,
            password: TestPassword);

        // Assert
        Assert.Equal(properties1, properties2);
    }

    [Fact]
    public void ConnectionProperties_WhenInfobasePathIsNull_ThrowsArgumentNullException() =>
        Assert.Throws<ArgumentNullException>(() => new ConnectionProperties(infobasePath: null, username: TestUsername, password: TestPassword));

    [Fact]
    public void ConnectionProperties_WhenTestUsernameIsNull_ThrowsArgumentNullException() =>
        Assert.Throws<ArgumentNullException>(() => new ConnectionProperties(infobasePath: TestInfobasePath, username: null, password: TestPassword));

    [Fact]
    public void ConnectionProperties_WhenPasswordIsNull_ShouldCreate() =>
        Assert.NotNull(() => new ConnectionProperties(infobasePath: TestInfobasePath, username: TestUsername, password: null));
}
