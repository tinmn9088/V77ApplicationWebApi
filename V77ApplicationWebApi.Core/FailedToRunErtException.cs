using System;

namespace V77ApplicationWebApi.Core;

public class FailedToRunErtException(string message)
    : Exception($"Failed to run ERT: {message}")
{
}
