using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Moq;
using V77ApplicationWebApi.Core;
using Xunit;
using Xunit.Abstractions;

namespace V77ApplicationWebApi.Infrastructure.UnitTests;

public class ComV77ApplicationConnectionFactoryTests
{
    private readonly ComV77ApplicationConnectionFactory _connectionFactory;

    private readonly Mock<IInstanceFactory> _instanceFactoryMock;

    private readonly Mock<IMemberInvoker> _memberInvokerMock;

    private readonly Mock<ILogger<ComV77ApplicationConnection>> _connectionLoggerMock;

    public ComV77ApplicationConnectionFactoryTests(ITestOutputHelper output)
    {
        _instanceFactoryMock = new();
        _memberInvokerMock = new();
        _connectionLoggerMock = new();

        _connectionFactory = new(
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            connectionLogger: _connectionLoggerMock.Object);
    }

    private static ConnectionProperties TestConnectionProperties => new(
        infobasePath: @"D:\TestInfobase",
        username: "TestUser",
        password: "TestPassword");

    [Fact]
    public async Task GetConnectionAsync_ShouldCreateAndSaveConnection()
    {
        // Arrange
        ConnectionProperties properties = TestConnectionProperties;

        // Act
        ComV77ApplicationConnection connection1 = await _connectionFactory.GetConnectionAsync(properties, CancellationToken.None) as ComV77ApplicationConnection;
        ComV77ApplicationConnection connection2 = await _connectionFactory.GetConnectionAsync(properties, CancellationToken.None) as ComV77ApplicationConnection;

        Assert.NotNull(connection1);
        Assert.NotNull(connection2);
        Assert.True(ReferenceEquals(connection1, connection2));
        Assert.Equal(1, _connectionFactory.InstancesCount);
    }
}
