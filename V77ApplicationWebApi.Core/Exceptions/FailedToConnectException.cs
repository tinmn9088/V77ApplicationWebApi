using System;

namespace V77ApplicationWebApi.Core.Exceptions;

public class FailedToConnectException(string infobasePath, Exception innerException)
    : Exception($"Failed to connect to infobase '{infobasePath}'", innerException)
{
}
