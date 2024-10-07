using System;

namespace V77ApplicationWebApi.Infrastructure.Exceptions;

public class FailedToGetTypeFromProgIDException(string progID, Exception innerException)
    : Exception($"Failed to get type from ProgID '{progID}'", innerException)
{
}
