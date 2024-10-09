using Microsoft.Extensions.Logging;
using Moq;
using V77ApplicationWebApi.Infrastructure.Exceptions;
using V77ApplicationWebApi.Infrastructure.UnitTests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace V77ApplicationWebApi.Infrastructure.UnitTests;

public class MemberInvokerTests
{
    private readonly MemberInvoker _memberInvoker;

    private readonly Mock<ILogger<MemberInvoker>> _loggerMock;

    public MemberInvokerTests(ITestOutputHelper output)
    {
        _loggerMock = new();
        _loggerMock.RedirectToStandartOutput(output);
        _memberInvoker = new(_loggerMock.Object);
    }

    [Fact]
    public void InvokePublicMethodByName_ShouldInvokeMethodAndReturnResult()
    {
        // Arrange
        object target = new();
        int expectedResult = target.GetHashCode();

        // Act
        int actualResult = (int)_memberInvoker.InvokePublicMethodByName(target, methodName: "GetHashCode", args: null);

        // Assert
        Assert.Equal(expectedResult, actualResult);
    }

    [Fact]
    public void InvokePublicMethodByName_WhenMethodNotFound_ThrowsFailedToInvokeMemberException() =>
        Assert.Throws<FailedToInvokeMemberException>(() => _memberInvoker.InvokePublicMethodByName(new object(), "NonExistingMethodName", null));

    [Fact]
    public void GetPropertyValueByName_ShouldReturnPropertyValue()
    {
        // Arrange
        int expectedPropertyValue = 42;
        ClassWithProperty target = new(expectedPropertyValue);

        // Act
        int actualPropertyValue = (int)_memberInvoker.GetPropertyValueByName(target, propertyName: "Property");

        // Assert
        Assert.Equal(expectedPropertyValue, actualPropertyValue);
    }

    [Fact]
    public void GetPropertyValueByName_WhenPropertyNotFound_ThrowsFailedToInvokeMemberException() =>
        Assert.Throws<FailedToInvokeMemberException>(() => _memberInvoker.GetPropertyValueByName(new object(), "NonExistingPropertyName"));

    private class ClassWithProperty(int propertyValue)
    {
        public int Property { get; } = propertyValue;
    }
}
