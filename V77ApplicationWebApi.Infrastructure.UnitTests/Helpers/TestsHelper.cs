using System;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit.Abstractions;

namespace V77ApplicationWebApi.Infrastructure.UnitTests.Helpers;

internal static class TestsHelper
{
    public static void RedirectToStandartOutput<T>(this Mock<ILogger<T>> loggerMock, ITestOutputHelper output) =>
        _ = loggerMock
            .Setup(l => l.Log(
                It.IsAny<LogLevel>(), // Entry will be written on this level
                It.IsAny<EventId>(), // Id of the event
                It.Is<It.IsAnyType>((_, _) => true), // The entry to be written
                It.IsAny<Exception>(), // The exception related to this entry
                It.IsAny<Func<It.IsAnyType, Exception?, string>>())) // Function to create a message
            .Callback((IInvocation i) => output.WriteLine($"{ExtractLogLevel(i)}: {ExtractLogMessage(i)}"));

    private static string ExtractLogLevel(IInvocation invovation) => invovation.Arguments[0].ToString().ToUpper();

    private static string ExtractLogMessage(IInvocation invovation) => invovation.Arguments[2].ToString();
}
