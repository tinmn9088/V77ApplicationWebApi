using V77ApplicationWebApi.Infrastructure.Exceptions;

namespace V77ApplicationWebApi.Infrastructure;

internal interface IMemberInvoker
{
    /// <summary>
    /// Invokes a public method of the <paramref name="target"/> by <paramref name="methodName"/>.
    /// </summary>
    /// <param name="target">The object on which to invoke the method.</param>
    /// <param name="methodName">The string containing the name of the method to invoke.</param>
    /// <param name="args">An array containing the arguments to pass to the method.</param>
    /// <returns>An <see cref="object"/> representing the return value of the invoked member.</returns>
    /// <exception cref="FailedToInvokeMemberException">If failed to invoke method.</exception>
    object? InvokePublicMethodByName(object target, string methodName, object[]? args);

    /// <summary>
    /// Get property value of the <paramref name="target"/> by <paramref name="propertyName"/>.
    /// </summary>
    /// <param name="target">The object of which to get the property value.</param>
    /// <param name="propertyName">The string containing the name of the property.</param>
    /// <returns>An <see cref="object"/> representing the value of the property.</returns>
    /// <exception cref="FailedToInvokeMemberException">If failed to invoke method.</exception>
    object? GetPropertyValueByName(object target, string propertyName);
}
