using Microsoft.Extensions.Logging;
using static V77ApplicationWebApi.Infrastructure.Helpers.MessageHelper;

namespace V77ApplicationWebApi.Infrastructure.Logging;

internal static class LoggerExtensions
{
    public static void LogTryingConnect(this ILogger logger, string infobasePath) =>
        logger.LogTrace("Trying to connect to infobase '{InfobasePath}'", infobasePath);

    public static void LogInitializingConnection(this ILogger logger, string infobasePath) =>
        logger.LogTrace("Initializing connection to infobase '{InfobasePath}'", infobasePath);

    public static void LogAlreadyConnected(this ILogger logger, string infobasePath) =>
        logger.LogTrace("Already connected to infobase '{InfobasePath}'", infobasePath);

    public static void LogInvokingMember(this ILogger logger, object target, string memberName, object[]? args) =>
        logger.LogTrace("Invoke '{MemberName}' on {Target} with args: {ArgsString}", memberName, target, BuildArgsString(args));
}
