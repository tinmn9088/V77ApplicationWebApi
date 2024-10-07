using System;
using System.Reflection;
using static System.Reflection.BindingFlags;

namespace V77ApplicationWebApi.Infrastructure;

internal class MemberInvoker : IMemberInvoker
{
    public object? InvokePublicMethodByName(object target, string methodName, object[] args) =>
        InvokeMember(
            target: target,
            memberName: methodName,
            attributes: Public | InvokeMethod,
            args: args);

    public object? GetPropertyValueByName(object target, string propertyName) =>
        InvokeMember(
            target: target,
            memberName: propertyName,
            attributes: GetProperty);

    private static object? InvokeMember(object target, string memberName, BindingFlags attributes, object[]? args = null)
    {
        Type targetType = target.GetType();

        return targetType.InvokeMember(
            target: target,
            name: memberName,
            invokeAttr: attributes,
            args: args,
            binder: null);
    }
}
