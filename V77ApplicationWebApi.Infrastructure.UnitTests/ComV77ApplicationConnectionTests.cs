using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Moq;
using V77ApplicationWebApi.Core;
using Xunit;
using Xunit.Abstractions;

namespace V77ApplicationWebApi.Infrastructure.UnitTests;

public class ComV77ApplicationConnectionTests
{
    private readonly ComV77ApplicationConnection _connection;

    private readonly Mock<IInstanceFactory> _instanceFactoryMock;

    private readonly Mock<IMemberInvoker> _memberInvokerMock;

    private readonly Mock<ILogger<ComV77ApplicationConnection>> _loggerMock;

    private readonly ConnectionProperties _testConnectionProperties;

    public ComV77ApplicationConnectionTests(ITestOutputHelper output)
    {
        _testConnectionProperties = new(
            infobasePath: @"D:\TestInfobase",
            username: "TestUser",
            password: "TestPassword");

        _instanceFactoryMock = new();
        _memberInvokerMock = new();
        _loggerMock = new();
        _loggerMock.RedirectToStandartOutput(output);

        _connection = new(
            properties: _testConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
    }

    [Fact]
    public async Task ConnectAsync_ShouldCreateAndInitializeComObject()
    {
        // Arrange
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = true;

        // Setup
        _ = _instanceFactoryMock
            .Setup(f => f.GetTypeFromProgID(It.IsAny<string>()))
            .Returns(fakeComObjectType);
        _ = _instanceFactoryMock
            .Setup(f => f.CreateInstance(It.IsAny<Type>()))
            .Returns(fakeComObject);
        _ = _memberInvokerMock
            .Setup(i => i.GetPropertyValueByName(
                It.IsAny<It.IsAnyType>(),
                It.IsAny<string>()))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.IsAny<string>(),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);

        // Act
        await _connection.ConnectAsync(CancellationToken.None);

        // Verify
        _instanceFactoryMock
            .Verify(f => f.GetTypeFromProgID(It.Is<string>(p => p == ComV77ApplicationConnection.ComObjectTypeName)), Times.Once);
        _instanceFactoryMock
            .Verify(f => f.CreateInstance(It.Is<Type>(t => t == fakeComObjectType)), Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.GetPropertyValueByName(
                    It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                    It.Is<string>(n => n == "RMTrade")),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                    It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                    It.Is<string>(n => n == "Initialize"),
                    It.Is<object[]>(a => a.Length == 3)),
                Times.Once);
    }
}
