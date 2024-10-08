using System;

namespace V77ApplicationWebApi.Infrastructure.Exceptions;

public class FailedToInitializeException(string infobasePath)
    : Exception($"Failed to initialize connection to infobase '{infobasePath}'")
{
}
