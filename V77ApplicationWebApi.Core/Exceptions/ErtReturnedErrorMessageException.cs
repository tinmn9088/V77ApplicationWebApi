using System;

namespace V77ApplicationWebApi.Core.Exceptions;

public class ErtReturnedErrorMessageException(string infobasePath, string message)
    : Exception($"ERT running at '{infobasePath}' returned error message: '{message}'")
{
}
