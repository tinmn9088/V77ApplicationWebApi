using System;

namespace V77ApplicationWebApi.Core.Exceptions;

public class FailedToRunErtException(string infobasePath, Exception innerException)
    : Exception($"Failed to run ERT at infobase '{infobasePath}'", innerException)
{
}
