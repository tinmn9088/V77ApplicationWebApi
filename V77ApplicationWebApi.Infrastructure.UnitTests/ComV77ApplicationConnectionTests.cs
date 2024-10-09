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
using V77ApplicationWebApi.Infrastructure.UnitTests.TestsHelpers;
using Xunit;
using Xunit.Abstractions;

namespace V77ApplicationWebApi.Infrastructure.UnitTests;

public class ComV77ApplicationConnectionTests
{
    private const string DefaultErtRelativePath = @"ExtForms\Test\Run.ert";

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
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();

        // Setup
        _instanceFactoryMock.SetupToConnect(comObject: fakeComObject, comObjectType: fakeComObjectType);
        _memberInvokerMock.SetupToInitializeConnection(rmtrade: 42, initializeResult: true);

        // Act
        await connection.ConnectAsync(CancellationToken.None);

        // Verify
        _instanceFactoryMock.VerifyGetTypeFromProgID(progID: ComV77ApplicationConnection.ComObjectTypeName, times: Times.Once);
        _instanceFactoryMock.VerifyCreateInstance(type: fakeComObjectType, times: Times.Once);
        _memberInvokerMock.VerifyGetPropertyValueByName(target: fakeComObject, propertyName: "RMTrade", times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(target: fakeComObject, methodName: "Initialize", verifyArgs: a => a.Length == 3, times: Times.Once);
        _instanceFactoryMock.VerifyNoOtherCalls();
        _memberInvokerMock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task ConnectAsync_WhenCalledMultipleTimes_ShouldCreateAndInitializeComObjectOnce()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();

        // Setup
        _instanceFactoryMock.SetupToConnect(comObject: fakeComObject, comObjectType: fakeComObjectType);
        _memberInvokerMock.SetupToInitializeConnection(rmtrade: 42, initializeResult: true);

        // Act
        for (int i = 0; i < 3; i++)
        {
            await connection.ConnectAsync(CancellationToken.None);
        }

        // Verify
        _instanceFactoryMock.VerifyGetTypeFromProgID(progID: ComV77ApplicationConnection.ComObjectTypeName, times: Times.Once);
        _instanceFactoryMock.VerifyCreateInstance(type: fakeComObjectType, times: Times.Once);
        _memberInvokerMock.VerifyGetPropertyValueByName(target: fakeComObject, propertyName: "RMTrade", times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(target: fakeComObject, methodName: "Initialize", verifyArgs: a => a.Length == 3, times: Times.Once);
        _instanceFactoryMock.VerifyNoOtherCalls();
        _memberInvokerMock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task ConnectAsync_WhenOperationIsCancelled_ThrowsOperationCanceledExceptionAndNotIncrementsComObjectErrorsCount()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        Type fakeComObjectType = typeof(object);
        using CancellationTokenSource tokenSource = new();

        // Setup
        _ = _instanceFactoryMock
            .SetupGetTypeFromProgID()
            .Callback((string _) => tokenSource.Cancel())
            .Returns(fakeComObjectType);

        // Act
        _ = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => connection.ConnectAsync(tokenSource.Token).AsTask());

        // Verify
        _memberInvokerMock.VerifyNoOtherCalls();

        // Assert
        Assert.Equal(0, connection.ComObjectErrorsCount);
    }

    [Fact]
    public async Task ConnectAsync_WhenCannotCreateType_ThrowsFailedToConnectExceptionAndIncrementsComObjectErrorsCount()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();

        // Setup
        _ = _instanceFactoryMock
            .SetupGetTypeFromProgID()
            .Throws<Exception>();

        // Act
        _ = await Assert.ThrowsAnyAsync<FailedToConnectException>(() => connection.ConnectAsync(CancellationToken.None).AsTask());

        // Verify
        _instanceFactoryMock.VerifyGetTypeFromProgID(progID: ComV77ApplicationConnection.ComObjectTypeName, times: Times.Once);
        _instanceFactoryMock.VerifyNoOtherCalls();
        _memberInvokerMock.VerifyNoOtherCalls();

        // Assert
        Assert.Equal(1, connection.ComObjectErrorsCount);
    }

    [Fact]
    public async Task ConnectAsync_WhenTooManyErrors_ThrowsErrorsCountExceededException()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();

        // Setup
        _ = _instanceFactoryMock
            .Setup(f => f.GetTypeFromProgID(It.IsAny<string>()))
            .Throws<Exception>();

        // Act, Assert
        for (int i = 0; i < ComV77ApplicationConnection.MaxComObjectErrorsCount; i++)
        {
            ValueTask connectTask = connection.ConnectAsync(CancellationToken.None);
            _ = await Assert.ThrowsAnyAsync<Exception>(connectTask.AsTask);
        }

        Assert.Equal(ComV77ApplicationConnection.MaxComObjectErrorsCount, connection.ComObjectErrorsCount);

        _ = await Assert.ThrowsAsync<ErrorsCountExceededException>(() => connection.ConnectAsync(CancellationToken.None).AsTask());

        // ComObjectErrorsCount must not exceed MaxComObjectErrorsCount after sequential calls
        Assert.Equal(ComV77ApplicationConnection.MaxComObjectErrorsCount, connection.ComObjectErrorsCount);
    }

    [Fact]
    public async Task ConnectAsync_WhenFailsToInvokeMember_ThrowsFailedToConnectExceptionAndIncrementsComObjectErrorsCount()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject, fakeComObjectType);
        _ = _memberInvokerMock
            .SetupGetPropertyValueByName()
            .Throws<FailedToInvokeMemberException>(() => new(
                target: fakeComObject,
                memberName: "ErrorProneMemberName",
                args: null,
                innerException: null));

        // Act
        _ = await Assert.ThrowsAsync<FailedToConnectException>(() => connection.ConnectAsync(CancellationToken.None).AsTask());

        // Verify
        _instanceFactoryMock.VerifyGetTypeFromProgID(progID: ComV77ApplicationConnection.ComObjectTypeName, times: Times.Once);
        _instanceFactoryMock.VerifyCreateInstance(type: fakeComObjectType, times: Times.Once);
        _instanceFactoryMock.VerifyNoOtherCalls();

        // Assert
        Assert.Equal(1, connection.ComObjectErrorsCount);
    }

    [Fact]
    public async Task ConnectAsync_WhenExceedesInitializeTimeout_ThrowsFailedToConnectException()
    {
        // Arrange
        ConnectionProperties propertiesWithZeroTimeout = new(
            infobasePath: DefaultConnectionProperties.InfobasePath,
            username: DefaultConnectionProperties.Username,
            password: DefaultConnectionProperties.Password,
            initializeTimeout: TimeSpan.Zero);
        ComV77ApplicationConnection connection = new(
            properties: propertiesWithZeroTimeout,
            instanceFactory: _instanceFactoryMock.Object,
            memberInvoker: _memberInvokerMock.Object,
            logger: _loggerMock.Object);
        object fakeComObject = new();
        Type fakeComObjectType = fakeComObject.GetType();

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject, fakeComObjectType);
        _ = _memberInvokerMock
            .SetupGetPropertyValueByName()
            .Callback((object _, string _) => Task.Delay(Timeout.InfiniteTimeSpan).Wait());

        // Act
        _ = await Assert.ThrowsAsync<FailedToConnectException>(() => connection.ConnectAsync(CancellationToken.None).AsTask());

        // Verify
        _instanceFactoryMock.VerifyGetTypeFromProgID(progID: ComV77ApplicationConnection.ComObjectTypeName, times: Times.Once);
        _instanceFactoryMock.VerifyCreateInstance(type: fakeComObjectType, times: Times.Once);
        _memberInvokerMock.VerifyGetPropertyValueByName(target: fakeComObject, propertyName: "RMTrade", times: Times.Once);
        _instanceFactoryMock.VerifyNoOtherCalls();
        _memberInvokerMock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task RunErtAsync_ShouldRunErt()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _memberInvokerMock.SetupToRunErt(out object fakeContextValueList);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        await connection.RunErtAsync(DefaultErtRelativePath, CancellationToken.None);

        // Verify
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "CreateObject",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == "ValueList",
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "OpenForm",
            verifyArgs: a =>
                a.Length == 3
                && (a[0] as string) == "Report"
                && ReferenceEquals(a[1], fakeContextValueList)
                && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, DefaultErtRelativePath),
            times: Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenErtContextPassed_ShouldRunErt()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        Dictionary<string, string> ertContext = new()
        {
            { "Key1", "Value1" },
            { "Key2", "Value2" },
        };

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _memberInvokerMock.SetupToRunErt(out object fakeContextValueList);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        await connection.RunErtAsync(DefaultErtRelativePath, ertContext, CancellationToken.None);

        // Verify
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "CreateObject",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == "ValueList",
            times: Times.Once);

        foreach (KeyValuePair<string, string> ertContextEntry in ertContext)
        {
            _memberInvokerMock.VerifyInvokePublicMethodByName(
                target: fakeContextValueList,
                methodName: "AddValue",
                verifyArgs: a =>
                    a.Length == 2
                    && (a[0] as string) == ertContextEntry.Value
                    && (a[1] as string) == ertContextEntry.Key,
                times: Times.Once);
        }

        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "OpenForm",
            verifyArgs: a =>
                a.Length == 3
                && (a[0] as string) == "Report"
                && ReferenceEquals(a[1], fakeContextValueList)
                && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, DefaultErtRelativePath),
            times: Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenResultNamePassed_ShouldRunErtAndReturnResult()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        string resultName = "TestErtResultName";
        string expectedResult = "TestErtResult";

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _memberInvokerMock.SetupToRunErt(out object fakeContextValueList, resultName, expectedResult);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        string actualResult = await connection.RunErtAsync(DefaultErtRelativePath, ertContext: null, resultName, CancellationToken.None);

        // Assert
        Assert.Equal(expectedResult, actualResult);

        // Verify
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "CreateObject",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == "ValueList",
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "OpenForm",
            verifyArgs: a =>
                a.Length == 3
                && (a[0] as string) == "Report"
                && ReferenceEquals(a[1], fakeContextValueList)
                && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, DefaultErtRelativePath),
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeContextValueList,
            methodName: "Get",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == resultName,
            times: Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenResultNamePassedAndResultIsNull_ShouldRunErtAndReturnNull()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        string resultName = "TestErtNullResultName";

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _memberInvokerMock.SetupToRunErt(out object fakeContextValueList, resultName, result: null);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        string result = await connection.RunErtAsync(DefaultErtRelativePath, ertContext: null, resultName, CancellationToken.None);

        // Assert
        Assert.Null(result);

        // Verify
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "CreateObject",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == "ValueList",
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "OpenForm",
            verifyArgs: a =>
                a.Length == 3
                && (a[0] as string) == "Report"
                && ReferenceEquals(a[1], fakeContextValueList)
                && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, DefaultErtRelativePath),
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeContextValueList,
            methodName: "Get",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == resultName,
            times: Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenErrorMessageNamePassedAndErrorMessageIsNullAfterRun_ShouldRunErt()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        string errorMessageName = "TestErtErrorMessageName";

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _memberInvokerMock.SetupToRunErt(out object fakeContextValueList, resultName: null, result: null, errorMessageName, errorMessage: null);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        _ = await connection.RunErtAsync(DefaultErtRelativePath, ertContext: null, resultName: null, errorMessageName, CancellationToken.None);

        // Verify
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "CreateObject",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == "ValueList",
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "OpenForm",
            verifyArgs: a =>
                a.Length == 3
                && (a[0] as string) == "Report"
                && ReferenceEquals(a[1], fakeContextValueList)
                && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, DefaultErtRelativePath),
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeContextValueList,
            methodName: "Get",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == errorMessageName,
            times: Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenErrorMessageNameProvidedAndErrorMessageIsNotNullAfterRun_ThrowsFailedToRunErtException()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        string errorMessageName = "TestErtErrorMessageName";
        string errorMessage = "Error code - 1";

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _memberInvokerMock.SetupToRunErt(out object fakeContextValueList, resultName: null, result: null, errorMessageName, errorMessage);

        // Act
        await connection.ConnectAsync(CancellationToken.None);
        _ = await Assert.ThrowsAsync<FailedToRunErtException>(() => connection.RunErtAsync(DefaultErtRelativePath, ertContext: null, resultName: null, errorMessageName, CancellationToken.None).AsTask());

        // Verify
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "CreateObject",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == "ValueList",
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeComObject,
            methodName: "OpenForm",
            verifyArgs: a =>
                a.Length == 3
                && (a[0] as string) == "Report"
                && ReferenceEquals(a[1], fakeContextValueList)
                && (a[2] as string) == Path.Combine(connection.Properties.InfobasePath, DefaultErtRelativePath),
            times: Times.Once);
        _memberInvokerMock.VerifyInvokePublicMethodByName(
            target: fakeContextValueList,
            methodName: "Get",
            verifyArgs: a => a.Length == 1 && (a[0] as string) == errorMessageName,
            times: Times.Once);
    }

    [Fact]
    public async Task RunErtAsync_WhenOperationIsCancelled_ThrowsOperationCanceledExceptionAndNotIncrementsComObjectErrorsCount()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();
        object fakeContextValueList = new();
        using CancellationTokenSource tokenSource = new();

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _memberInvokerMock.SetupToInitializeConnection(rmtrade: 42, initializeResult: true);
        _ = _memberInvokerMock
            .SetupInvokePublicMethodByName(methodName: "CreateObject")
            .Callback((object _, string _, object[] _) => tokenSource.Cancel())
            .Returns(fakeContextValueList);

        // Act, Assert
        await connection.ConnectAsync(CancellationToken.None);
        _ = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => connection.RunErtAsync(DefaultErtRelativePath, tokenSource.Token).AsTask());
    }

    [Fact]
    public async Task RunErtAsyncsync_WhenTooManyErrors_ThrowsErrorsCountExceededException()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _memberInvokerMock.SetupToInitializeConnection(rmtrade: 42, initializeResult: true);
        _ = _memberInvokerMock
            .SetupInvokePublicMethodByName(methodName: "CreateObject")
            .Throws<Exception>();

        // Act
        await connection.ConnectAsync(CancellationToken.None);

        for (int i = 0; i < ComV77ApplicationConnection.MaxComObjectErrorsCount; i++)
        {
            ValueTask runErtTask = connection.RunErtAsync(DefaultErtRelativePath, CancellationToken.None);
            _ = await Assert.ThrowsAnyAsync<Exception>(runErtTask.AsTask);
        }

        _ = await Assert.ThrowsAsync<ErrorsCountExceededException>(() => connection.RunErtAsync(DefaultErtRelativePath, CancellationToken.None).AsTask());

        // Assert
        Assert.Equal(ComV77ApplicationConnection.MaxComObjectErrorsCount, connection.ComObjectErrorsCount);
    }

    [Fact]
    public async Task RunErtAsyncsync_WhenComObjectNotCreated_ThrowsFailedToRunErtExceptionAndNotIncrementComObjectErrorsCount()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();

        // Act (missing ConnectAsync)
        _ = await Assert.ThrowsAsync<FailedToRunErtException>(() => connection.RunErtAsync(DefaultErtRelativePath, CancellationToken.None).AsTask());

        // Assert
        Assert.Equal(0, connection.ComObjectErrorsCount);
    }

    [Fact]
    public async Task RunErtAsyncsync_WhenIsNotInitialized_ThrowsFailedToRunErtExceptionAndNotIncrementComObjectErrorsCount()
    {
        // Arrange
        ComV77ApplicationConnection connection = CreateDefaultConnection();
        object fakeComObject = new();

        // Setup
        _instanceFactoryMock.SetupToConnect(fakeComObject);
        _ = _memberInvokerMock
            .SetupGetPropertyValueByName(propertyName: "RMTrade")
            .Returns(42);
        _ = _memberInvokerMock
            .SetupInvokePublicMethodByName(methodName: "Initialize")
            .Returns(false);

        // Act
        ValueTask connectTask = connection.ConnectAsync(CancellationToken.None);
        _ = await Assert.ThrowsAsync<FailedToConnectException>(connectTask.AsTask);
        ValueTask runErtTask = connection.RunErtAsync(DefaultErtRelativePath, CancellationToken.None);
        _ = await Assert.ThrowsAsync<FailedToRunErtException>(runErtTask.AsTask);

        // Assert
        Assert.Equal(1, connection.ComObjectErrorsCount);
    }

    private ComV77ApplicationConnection CreateDefaultConnection() => new(
        properties: DefaultConnectionProperties,
        instanceFactory: _instanceFactoryMock.Object,
        memberInvoker: _memberInvokerMock.Object,
        logger: _loggerMock.Object);
}
