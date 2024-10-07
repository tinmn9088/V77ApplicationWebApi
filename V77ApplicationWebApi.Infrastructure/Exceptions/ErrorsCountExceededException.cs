using System;

namespace V77ApplicationWebApi.Infrastructure.Exceptions;

public class ErrorsCountExceededException(int errorsCount)
    : Exception($"Too many errors: {errorsCount}")
{
}
