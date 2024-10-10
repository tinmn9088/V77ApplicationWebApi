using System;
using System.Reflection;
using Microsoft.Extensions.Logging;
using V77ApplicationWebApi.Infrastructure.Exceptions;
using V77ApplicationWebApi.Infrastructure.Logging;
using static System.Reflection.BindingFlags;

namespace V77ApplicationWebApi.Infrastructure;

internal sealed class MemberInvoker(ILogger<MemberInvoker> logger)
    : IMemberInvoker
{
    private readonly ILogger<MemberInvoker> _logger = logger;

    public object? InvokePublicMethodByName(object target, string methodName, object[] args) =>
        InvokeMember(
            target: target,
            memberName: methodName,
            attributes: Public | Instance | InvokeMethod,
            args: args);

    public object? GetPropertyValueByName(object target, string propertyName) =>
        InvokeMember(
            target: target,
            memberName: propertyName,
            attributes: Public | Instance | GetProperty);

    private object? InvokeMember(object target, string memberName, BindingFlags attributes, object[]? args = null)
    {
        Type targetType = target.GetType();

        _logger.LogInvokingMember(target, memberName, args);

        try
        {
            return targetType.InvokeMember(
                target: target,
                name: memberName,
                invokeAttr: attributes,
                args: args,
                binder: null);
        }
        catch (Exception ex)
        {
            throw new FailedToInvokeMemberException(target, memberName, args, ex);
        }
    }
}
