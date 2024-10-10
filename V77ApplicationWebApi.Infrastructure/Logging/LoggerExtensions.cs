using System;
using Microsoft.Extensions.Logging;
using static V77ApplicationWebApi.Infrastructure.Helpers.MessageHelper;

namespace V77ApplicationWebApi.Infrastructure.Logging;

internal static class LoggerExtensions
{
    public static void LogTryingConnect(this ILogger logger, string infobasePath) =>
        logger.LogTrace("Trying to connect to infobase '{InfobasePath}'", infobasePath);

    public static void LogInitializingConnection(this ILogger logger, string infobasePath) =>
        logger.LogInformation("Initializing connection to infobase '{InfobasePath}'", infobasePath);

    public static void LogAlreadyConnected(this ILogger logger, string infobasePath) =>
        logger.LogTrace("Already connected to infobase '{InfobasePath}'", infobasePath);

    public static void LogRunningErt(this ILogger logger, string infobasePath, string ertRelativePath) =>
        logger.LogInformation("Running ERT '{ErtRelativePath}' at infobase '{InfobasePath}'", ertRelativePath, infobasePath);

    public static void LogInvokingMember(this ILogger logger, object target, string memberName, object[]? args) =>
        logger.LogTrace("Invoke '{MemberName}' on {Target} with args: {ArgsString}", memberName, target, BuildArgsString(args));

    public static void LogConnectionDisposing(this ILogger logger, string infobasePath, TimeSpan disposeTimeout) =>
        logger.LogTrace("Connection to infobase '{InfobasePath}' was inactive for {DisposeTimeout}. Disposing ...", infobasePath, disposeTimeout);

    public static void LogErrorOnBeforeDisposeConnection(this ILogger logger, Exception ex, string infobasePath) =>
        logger.LogError(ex, "Error in before dispose callback for connection to infobase '{InfobasePath}'", infobasePath);

    public static void LogErrorOnAfterDisposeConnection(this ILogger logger, Exception ex, string infobasePath) =>
        logger.LogError(ex, "Error in after dispose callback for connection to infobase '{InfobasePath}'", infobasePath);

    public static void LogConnectionDisposed(this ILogger logger, string infobasePath) =>
        logger.LogTrace("Connection to infobase '{InfobasePath}' disposed", infobasePath);
}
