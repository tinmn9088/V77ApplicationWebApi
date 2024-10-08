using System;
using Moq;
using Moq.Language.Flow;

namespace V77ApplicationWebApi.Infrastructure.UnitTests;

internal static class MockHelper
{
    public static ISetup<IMemberInvoker, object> SetupGetPropertyValueByName(this Mock<IMemberInvoker> memberInvokerMock, string propertyName) =>
        memberInvokerMock
            .Setup(i => i.GetPropertyValueByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == propertyName)));

    public static ISetup<IMemberInvoker, object> SetupInvokePublicMethodByName(this Mock<IMemberInvoker> memberInvokerMock, string methodName, Func<object[], bool> detectArgs) =>
        memberInvokerMock
            .Setup(i => i.InvokePublicMethodByName(
                It.IsAny<It.IsAnyType>(),
                It.Is<string>(n => n == methodName),
                It.Is<object[]>(a => detectArgs(a))));
}
