using Microsoft.Extensions.Logging;

namespace V77ApplicationWebApi.Infrastructure.Logging;

internal static class LoggerExtensions
{
    public static void LogTryingConnect(this ILogger logger, string infobasePath) => logger.LogTrace("Trying to connect to '{InfobasePath}'", infobasePath);
}
