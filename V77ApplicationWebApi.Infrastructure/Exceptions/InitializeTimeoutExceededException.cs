using System;

namespace V77ApplicationWebApi.Infrastructure.Exceptions;

public class InitializeTimeoutExceededException(TimeSpan timeout)
    : Exception($"Initialize timeout exceeded ({timeout})")
{
}
