using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Moq;
using V77ApplicationWebApi.Core;
using V77ApplicationWebApi.Core.Exceptions;
using V77ApplicationWebApi.Infrastructure.Exceptions;
using Xunit;
using Xunit.Abstractions;

namespace V77ApplicationWebApi.Infrastructure.UnitTests;

public class ComV77ApplicationConnectionTests
{
    private readonly Mock<IInstanceFactory> _instanceFactoryMock;

    private readonly Mock<IMemberInvoker> _memberInvokerMock;

    private readonly Mock<ILogger<ComV77ApplicationConnection>> _loggerMock;

    public ComV77ApplicationConnectionTests(ITestOutputHelper output)
    {
        _instanceFactoryMock = new();
        _memberInvokerMock = new();
        _loggerMock = new();
        _loggerMock.RedirectToStandartOutput(output);
    }

    private static ConnectionProperties DefaultConnectionProperties => new(
        infobasePath: @"D:\TestInfobase",
        username: "TestUser",
        password: "TestPassword");

    [Fact]
    public async Task ConnectAsync_ShouldCreateAndInitializeComObject()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
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
        await connection.ConnectAsync(CancellationToken.None);

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

    [Fact]
    public async Task ConnectAsync_WhenCannotCreateType_ThrowsExceptionAndIncrementsErrorsCount()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);

        // Setup
        _ = _instanceFactoryMock
            .Setup(f => f.GetTypeFromProgID(It.IsAny<string>()))
            .Throws<Exception>();

        // Act
        _ = await Assert.ThrowsAnyAsync<Exception>(() => connection.ConnectAsync(CancellationToken.None).AsTask());

        // Verify
        _instanceFactoryMock
            .Verify(f => f.GetTypeFromProgID(It.Is<string>(p => p == ComV77ApplicationConnection.ComObjectTypeName)), Times.Once);
        _instanceFactoryMock
            .Verify(f => f.CreateInstance(It.IsAny<Type>()), Times.Never);
        _memberInvokerMock
            .Verify(
                i => i.GetPropertyValueByName(
                    It.IsAny<object>(),
                    It.IsAny<string>()),
                Times.Never);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                    It.IsAny<object>(),
                    It.IsAny<string>(),
                    It.IsAny<object[]>()),
                Times.Never);

        // Assert
        Assert.Equal(1, connection.ErrorsCount);
    }

    [Fact]
    public async Task ConnectAsync_WhenTooManyErrors_ThrowsErrorsCountExceededException()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);

        // Setup
        _ = _instanceFactoryMock
            .Setup(f => f.GetTypeFromProgID(It.IsAny<string>()))
            .Throws<Exception>();

        // Act
        for (int i = 0; i < ComV77ApplicationConnection.MaxErrorsCount; i++)
        {
            ValueTask connectTask = connection.ConnectAsync(CancellationToken.None);
            _ = await Assert.ThrowsAnyAsync<Exception>(connectTask.AsTask);
        }

        _ = await Assert.ThrowsAsync<FailedToConnectException>(() => connection.ConnectAsync(CancellationToken.None).AsTask());

        // Assert
        Assert.Equal(ComV77ApplicationConnection.MaxErrorsCount, connection.ErrorsCount);
    }

    [Fact]
    public async Task ConnectAsync_WhenErrorOnInvokeMember_ShouldNotIncrementsErrorsCount()
    {
        // Arrange
        ConnectionProperties properties = new(
            infobasePath: DefaultConnectionProperties.InfobasePath,
            username: DefaultConnectionProperties.Username,
            password: DefaultConnectionProperties.Password);
        ComV77ApplicationConnection connection = new(
            properties: properties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();

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
            .Throws<FailedToInvokeMemberException>(() => new(fakeComObject, "ErrorProneMemberName", args: null, innerException: null));

        // Act
        _ = await Assert.ThrowsAsync<FailedToConnectException>(() => connection.ConnectAsync(CancellationToken.None).AsTask());

        // Verify
        _instanceFactoryMock
            .Verify(f => f.GetTypeFromProgID(It.Is<string>(p => p == ComV77ApplicationConnection.ComObjectTypeName)), Times.Once);
        _instanceFactoryMock
            .Verify(f => f.CreateInstance(It.Is<Type>(t => t == fakeComObjectType)), Times.Once);

        // Assert
        Assert.Equal(1, connection.ErrorsCount);
    }

    [Fact]
    public async Task ConnectAsync_WhenExceedesInitializateTimeoutLimit_ThrowsInitializeTimeoutExceededException()
    {
        // Arrange
        ConnectionProperties properties = new(
            infobasePath: DefaultConnectionProperties.InfobasePath,
            username: DefaultConnectionProperties.Username,
            password: DefaultConnectionProperties.Password,
            initializeTimeout: TimeSpan.Zero);
        ComV77ApplicationConnection connection = new(
            properties: properties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();

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
            .Callback((object _, string _) => Task.Delay(Timeout.InfiniteTimeSpan).Wait());

        // Act
        _ = await Assert.ThrowsAsync<FailedToConnectException>(() => connection.ConnectAsync(CancellationToken.None).AsTask());

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
                    It.IsAny<object>(),
                    It.IsAny<string>(),
                    It.IsAny<object[]>()),
                Times.Never);
    }

    [Fact]
    public async Task RunErtAsync_ShouldRunErt()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = true;
        string testErtRelativePath = @"ExtForms\Test\Run.ert";
        object fakeContextValueList = new();

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
                It.Is<string>(n => n == "RMTrade")))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Initialize"),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")))
            .Returns(fakeContextValueList);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        await connection.RunErtAsync(testErtRelativePath, CancellationToken.None);

        // Verify
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "OpenForm"),
                It.Is<object[]>(a =>
                    a.Length == 3
                    && (a[0] as string) == "Report"
                    && ReferenceEquals(a[1], fakeContextValueList)
                    && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, testErtRelativePath))),
                Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenContextProvided_ShouldRunErt()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = true;
        string testErtRelativePath = @"ExtForms\Test\Run.ert";
        Dictionary<string, string> ertContext = new()
        {
            { "Key1", "Value1" },
            { "Key2", "Value2" },
        };
        object fakeContextValueList = new();

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
                It.Is<string>(n => n == "RMTrade")))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Initialize"),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")))
            .Returns(fakeContextValueList);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        await connection.RunErtAsync(testErtRelativePath, ertContext, CancellationToken.None);

        // Verify
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")),
                Times.Once);

        foreach (KeyValuePair<string, string> ertContextEntry in ertContext)
        {
            _memberInvokerMock
                .Verify(
                    i => i.InvokePublicMethodByName(
                    It.Is<object>(t => ReferenceEquals(t, fakeContextValueList)),
                    It.Is<string>(n => n == "AddValue"),
                    It.Is<object[]>(a =>
                        a.Length == 2
                        && (a[0] as string) == ertContextEntry.Value
                        && (a[1] as string) == ertContextEntry.Key)),
                    Times.Once);
        }

        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "OpenForm"),
                It.Is<object[]>(a =>
                    a.Length == 3
                    && (a[0] as string) == "Report"
                    && ReferenceEquals(a[1], fakeContextValueList)
                    && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, testErtRelativePath))),
                Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenResultNameProvided_ShouldRunErtAndReturnResult()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = true;
        string testErtRelativePath = @"ExtForms\Test\Run.ert";
        string resultName = "TestErtResultName";
        string fakeResult = "TestErtResult";
        object fakeContextValueList = new();

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
                It.Is<string>(n => n == "RMTrade")))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Initialize"),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")))
            .Returns(fakeContextValueList);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Get"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == resultName)))
            .Returns(fakeResult);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        string result = await connection.RunErtAsync(testErtRelativePath, ertContext: null, resultName, CancellationToken.None);

        // Assert
        Assert.Equal(fakeResult, result);

        // Verify
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "OpenForm"),
                It.Is<object[]>(a =>
                    a.Length == 3
                    && (a[0] as string) == "Report"
                    && ReferenceEquals(a[1], fakeContextValueList)
                    && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, testErtRelativePath))),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeContextValueList)),
                It.Is<string>(n => n == "Get"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == resultName)),
                Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenResultNameProvidedAndResultNull_ShouldRunErtAndReturnNull()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = true;
        string testErtRelativePath = @"ExtForms\Test\Run.ert";
        string resultName = "TestErtResultName";
        object fakeContextValueList = new();

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
                It.Is<string>(n => n == "RMTrade")))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Initialize"),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")))
            .Returns(fakeContextValueList);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Get"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == resultName)))
            .Returns(null);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        string result = await connection.RunErtAsync(testErtRelativePath, ertContext: null, resultName, CancellationToken.None);

        // Assert
        Assert.Null(result);

        // Verify
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "OpenForm"),
                It.Is<object[]>(a =>
                    a.Length == 3
                    && (a[0] as string) == "Report"
                    && ReferenceEquals(a[1], fakeContextValueList)
                    && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, testErtRelativePath))),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeContextValueList)),
                It.Is<string>(n => n == "Get"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == resultName)),
                Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenErrorMessageNameProvidedAndValueNullAfterRun_ShouldRunErt()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = true;
        string testErtRelativePath = @"ExtForms\Test\Run.ert";
        string errorMessageName = "TestErtErrorMessageName";
        object fakeContextValueList = new();

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
                It.Is<string>(n => n == "RMTrade")))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Initialize"),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")))
            .Returns(fakeContextValueList);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Get"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == errorMessageName)))
            .Returns(null);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        _ = await connection.RunErtAsync(testErtRelativePath, ertContext: null, resultName: null, errorMessageName, CancellationToken.None);

        // Verify
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "OpenForm"),
                It.Is<object[]>(a =>
                    a.Length == 3
                    && (a[0] as string) == "Report"
                    && ReferenceEquals(a[1], fakeContextValueList)
                    && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, testErtRelativePath))),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeContextValueList)),
                It.Is<string>(n => n == "Get"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == errorMessageName)),
                Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenErrorMessageNameProvidedAndValueNotNullAfterRun_ThrowsFailedToRunErtException()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = true;
        string testErtRelativePath = @"ExtForms\Test\Run.ert";
        string errorMessageName = "TestErtErrorMessageName";
        string errorMessage = "Error code - 1";
        object fakeContextValueList = new();

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
                It.Is<string>(n => n == "RMTrade")))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Initialize"),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")))
            .Returns(fakeContextValueList);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Get"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == errorMessageName)))
            .Returns(errorMessage);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        _ = await Assert.ThrowsAsync<FailedToRunErtException>(() => connection.RunErtAsync(testErtRelativePath, ertContext: null, resultName: null, errorMessageName, CancellationToken.None).AsTask());

        // Verify
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeComObject)),
                It.Is<string>(n => n == "OpenForm"),
                It.Is<object[]>(a =>
                    a.Length == 3
                    && (a[0] as string) == "Report"
                    && ReferenceEquals(a[1], fakeContextValueList)
                    && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, testErtRelativePath))),
                Times.Once);
        _memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                It.Is<object>(t => ReferenceEquals(t, fakeContextValueList)),
                It.Is<string>(n => n == "Get"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == errorMessageName)),
                Times.Once);
    }

    [Fact]
    public async Task RunErtAsyncsync_WhenTooManyErrors_ThrowsErrorsCountExceededException()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = true;
        string testErtRelativePath = @"ExtForms\Test\Run.ert";

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
                It.Is<string>(n => n == "RMTrade")))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Initialize"),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "CreateObject"),
                It.Is<object[]>(a => a.Length == 1 && (a[0] as string) == "ValueList")))
            .Throws<Exception>();
        await connection.ConnectAsync(CancellationToken.None);

        // Act
        for (int i = 0; i < ComV77ApplicationConnection.MaxErrorsCount; i++)
        {
            ValueTask runErtTask = connection.RunErtAsync(testErtRelativePath, CancellationToken.None);
            _ = await Assert.ThrowsAnyAsync<Exception>(runErtTask.AsTask);
        }

        _ = await Assert.ThrowsAsync<ErrorsCountExceededException>(() => connection.RunErtAsync(testErtRelativePath, CancellationToken.None).AsTask());

        // Assert
        Assert.Equal(ComV77ApplicationConnection.MaxErrorsCount, connection.ErrorsCount);
    }

    [Fact]
    public async Task RunErtAsyncsync_WhenComObjectNotCreated_ThrowsInvalidOperationExceptionAndNotIncrementErrorsCount()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        string testErtRelativePath = @"ExtForms\Test\Run.ert";

        // Act
        _ = await Assert.ThrowsAsync<InvalidOperationException>(() => connection.RunErtAsync(testErtRelativePath, CancellationToken.None).AsTask());

        // Assert
        Assert.Equal(0, connection.ErrorsCount);
    }

    [Fact]
    public async Task RunErtAsyncsync_WhenNotInitialized_ThrowsErrorsCountExceededException()
    {
        // Arrange
        ComV77ApplicationConnection connection = new(
            properties: DefaultConnectionProperties,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();
        int fakeRMTrade = 42;
        bool fakeInitializationResult = false;
        string testErtRelativePath = @"ExtForms\Test\Run.ert";

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
                It.Is<string>(n => n == "RMTrade")))
            .Returns(fakeRMTrade);
        _ = _memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == "Initialize"),
                It.IsAny<object[]?>()))
            .Returns(fakeInitializationResult);
        await connection.ConnectAsync(CancellationToken.None);

        // Act
        _ = await Assert.ThrowsAsync<InvalidOperationException>(() => connection.RunErtAsync(testErtRelativePath, CancellationToken.None).AsTask());

        // Assert
        Assert.Equal(0, connection.ErrorsCount);
    }
}
