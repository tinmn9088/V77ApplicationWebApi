using System;
using Moq;
using Moq.Language.Flow;

namespace V77ApplicationWebApi.Infrastructure.UnitTests.TestsHelpers;

internal static class MoqHelper
{
    public static ISetup<IInstanceFactory, Type> SetupGetTypeFromProgID(this Mock<IInstanceFactory> instanceFactoryMock) =>
        instanceFactoryMock.Setup(i => i.GetTypeFromProgID(It.IsAny<string>()));

    public static ISetup<IInstanceFactory, object> SetupCreateInstance(this Mock<IInstanceFactory> instanceFactoryMock) =>
        instanceFactoryMock.Setup(i => i.CreateInstance(It.IsAny<Type>()));

    public static void VerifyGetTypeFromProgID(this Mock<IInstanceFactory> instanceFactoryMock, string progID, Func<Times> times) =>
        instanceFactoryMock.Verify(f => f.GetTypeFromProgID(It.Is<string>(p => p == progID)), times);

    public static void VerifyCreateInstance(this Mock<IInstanceFactory> instanceFactoryMock, Type type, Func<Times> times) =>
        instanceFactoryMock.Verify(f => f.CreateInstance(It.Is<Type>(t => t == type)), times);

    /// <summary>
    /// Setup <see cref="IInstanceFactory.GetTypeFromProgID(string)"/> and <see cref="IInstanceFactory.CreateInstance(Type)"/> methods.
    /// </summary>
    /// <param name="instanceFactoryMock">Extensible instance.</param>
    /// <param name="comObject">Instance to return from <see cref="IInstanceFactory.CreateInstance(Type)"/>.</param>
    /// <param name="comObjectType">Type to return from <see cref="IInstanceFactory.GetTypeFromProgID(string)"/>.</param>
    public static void SetupToConnect(this Mock<IInstanceFactory> instanceFactoryMock, object comObject, Type? comObjectType = default)
    {
        _ = instanceFactoryMock
            .SetupGetTypeFromProgID()
            .Returns(comObjectType ?? comObject.GetType());
        _ = instanceFactoryMock
            .SetupCreateInstance()
            .Returns(comObject);
    }

    public static ISetup<IMemberInvoker, object> SetupGetPropertyValueByName(this Mock<IMemberInvoker> memberInvokerMock) =>
        memberInvokerMock
            .Setup(i => i.GetPropertyValueByName(
                It.IsAny<It.IsAnyType>(),
                It.IsAny<string>()));

    public static ISetup<IMemberInvoker, object> SetupGetPropertyValueByName(this Mock<IMemberInvoker> memberInvokerMock, string propertyName) =>
        memberInvokerMock
            .Setup(i => i.GetPropertyValueByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == propertyName)));

    public static ISetup<IMemberInvoker, object> SetupInvokePublicMethodByName(this Mock<IMemberInvoker> memberInvokerMock) =>
        memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.IsAny<string>(),
                It.IsAny<object[]>()));

    public static ISetup<IMemberInvoker, object> SetupInvokePublicMethodByName(this Mock<IMemberInvoker> memberInvokerMock, string methodName) =>
        memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == methodName),
                It.IsAny<object[]>()));

    public static ISetup<IMemberInvoker, object> SetupInvokePublicMethodByName(this Mock<IMemberInvoker> memberInvokerMock, string methodName, Func<object[], bool> verifyArgs) =>
        memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == methodName),
                It.Is<object[]>(a => verifyArgs(a))));

    public static void VerifyGetPropertyValueByName(this Mock<IMemberInvoker> memberInvokerMock, object target, string propertyName, Func<Times> times) =>
        memberInvokerMock
            .Verify(
                i => i.GetPropertyValueByName(
                    It.Is<object>(t => ReferenceEquals(t, target)),
                    It.Is<string>(n => n == propertyName)),
                times);

    public static void VerifyInvokePublicMethodByName(this Mock<IMemberInvoker> memberInvokerMock, object target, string methodName, Func<object[], bool> verifyArgs, Func<Times> times) =>
        memberInvokerMock
            .Verify(
                i => i.InvokePublicMethodByName(
                    It.Is<object>(t => ReferenceEquals(t, target)),
                    It.Is<string>(n => n == methodName),
                    It.Is<object[]>(a => verifyArgs(a))),
                times);

    /// <summary>
    /// Setup "RMTrade" and "Initialize" methods.
    /// </summary>
    /// <param name="memberInvokerMock">Extensible instance.</param>
    /// <param name="rmtrade">Value to return from <see cref="IMemberInvoker.GetPropertyValueByName(object, string)"/>.</param>
    /// <param name="initializeResult">Value to return from <see cref="IMemberInvoker.InvokePublicMethodByName(object, string, object[]?)"/>.</param>
    public static void SetupToInitializeConnection(this Mock<IMemberInvoker> memberInvokerMock, int rmtrade, bool initializeResult)
    {
        _ = memberInvokerMock
            .SetupGetPropertyValueByName(propertyName: "RMTrade")
            .Returns(rmtrade);
        _ = memberInvokerMock
            .SetupInvokePublicMethodByName(methodName: "Initialize")
            .Returns(initializeResult);
    }

    /// <summary>
    /// Setup "RMTrade", "Initialize", "CreateObject" methods.
    /// </summary>
    /// <param name="fakeContextValueList">Value to return after "CreateObject" call.</param>
    /// <inheritdoc cref="SetupToInitializeConnection(Mock{IMemberInvoker}, int, bool)"/>
    public static void SetupToRunErt(this Mock<IMemberInvoker> memberInvokerMock, out object fakeContextValueList)
    {
        memberInvokerMock.SetupToInitializeConnection(rmtrade: 42, initializeResult: true);

        fakeContextValueList = new();
        _ = memberInvokerMock
            .SetupInvokePublicMethodByName(methodName: "CreateObject")
            .Returns(fakeContextValueList);
    }

    /// <summary>
    /// Setup "RMTrade", "Initialize", "CreateObject", "Get" methods.
    /// </summary>
    /// <param name="resultName">Need to be passed to "Get" method to retrieve <paramref name="result"/>.</param>
    /// <param name="result">Value to return after "Get" call.</param>
    /// <inheritdoc cref="SetupToRunErt(Mock{IMemberInvoker}, out object)"/>
    public static void SetupToRunErt(this Mock<IMemberInvoker> memberInvokerMock, out object fakeContextValueList, string resultName, object? result)
    {
        memberInvokerMock.SetupToRunErt(out object contextValueList);

        fakeContextValueList = contextValueList;

        _ = memberInvokerMock
            .SetupInvokePublicMethodByName(
                methodName: "Get",
                verifyArgs: a => a.Length == 1 && (a[0] as string) == resultName)
            .Returns(result);
    }

    /// <summary>
    /// Setup "RMTrade", "Initialize", "CreateObject", "Get" methods (for result and for error message).
    /// </summary>
    /// <param name="resultName">Need to be passed to "Get" method to retrieve <paramref name="result"/>.</param>
    /// <param name="result">Value to return after "Get" call with <paramref name="resultName"/>.</param>
    /// <param name="errorMessageName">Need to be passed to "Get" method to retrieve <paramref name="errorMessage"/>.</param>
    /// <param name="errorMessage">Value to return after "Get" call with <paramref name="errorMessageName"/>.</param>
    /// <inheritdoc cref="SetupToRunErt(Mock{IMemberInvoker}, out object)"/>
    public static void SetupToRunErt(this Mock<IMemberInvoker> memberInvokerMock, out object fakeContextValueList, string resultName, object? result, string errorMessageName, object? errorMessage)
    {
        memberInvokerMock.SetupToRunErt(out object contextValueList, resultName, result);

        fakeContextValueList = contextValueList;

        _ = memberInvokerMock
            .SetupInvokePublicMethodByName(
                methodName: "Get",
                verifyArgs: a => a.Length == 1 && (a[0] as string) == errorMessageName)
            .Returns(errorMessage);
    }
}
