using System;
using V77ApplicationWebApi.Infrastructure.Exceptions;
using Xunit;

namespace V77ApplicationWebApi.Infrastructure.UnitTests;

public class InstanceFactoryTests
{
    private readonly InstanceFactory _instanceFactory;

    public InstanceFactoryTests() => _instanceFactory = new();

    [Fact]
    public void GetTypeFromProgID_ShouldReturnType()
    {
        // Arrange
        Type expected = typeof(object);

        // Act
        Type actual = _instanceFactory.GetTypeFromProgID("System.Object");

        // Assert
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void GetTypeFromProgID_WhenProgIDIsInvalid_ThrowsFailedToGetTypeFromProgIDException() =>
        Assert.Throws<FailedToGetTypeFromProgIDException>(() => _instanceFactory.GetTypeFromProgID("NonExistingProgID"));

    [Fact]
    public void CreateInstance_ShouldCreateInstance()
    {
        // Arrange
        Type expectedType = typeof(object);

        // Act
        object instance = _instanceFactory.CreateInstance(expectedType);

        // Assert
        Assert.IsType(expectedType, instance);
    }

    [Fact]
    public void CreateInstance_WhenFailsToCreateInstance_ThrowsFailedToCreateInstanceException() =>
        Assert.Throws<FailedToCreateInstanceException>(() => _instanceFactory.CreateInstance(typeof(ClassWithoutDefaultConstructor)));

    private class ClassWithoutDefaultConstructor(object value)
    {
        public object Value { get; } = value;
    }
}
