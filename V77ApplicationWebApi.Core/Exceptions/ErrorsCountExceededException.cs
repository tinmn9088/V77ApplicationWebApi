using System;

namespace V77ApplicationWebApi.Core.Exceptions;

public class ErrorsCountExceededException(int errorsCount)
    : Exception($"Too many errors: {errorsCount}")
{
}
