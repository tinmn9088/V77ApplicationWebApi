using System;

namespace V77ApplicationWebApi.Infrastructure.Exceptions;

public class FailedToCreateInstanceException(Type type, Exception innerException)
    : Exception($"Failed to create instance of type '{type}'", innerException)
{
}
