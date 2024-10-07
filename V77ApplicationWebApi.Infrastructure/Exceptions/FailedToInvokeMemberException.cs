using System;
using static V77ApplicationWebApi.Infrastructure.Helpers.MessageHelper;

namespace V77ApplicationWebApi.Infrastructure.Exceptions;

public class FailedToInvokeMemberException(object target, string memberName, object[]? args, Exception innerException)
    : Exception($"Failed to invoke member '{memberName}' on '{target}' with args: {BuildArgsString(args)}", innerException)
{
}
