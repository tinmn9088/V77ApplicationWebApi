using System;

namespace V77ApplicationWebApi.Core.Exceptions;

public class FailedToRunErtException(string infobasePath, string message)
    : Exception($"Failed to run ERT at infobase '{infobasePath}': {message}")
{
}
