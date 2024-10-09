using System;

namespace V77ApplicationWebApi.Core;

public record ConnectionProperties
{
    public ConnectionProperties(string infobasePath, string username, string? password, TimeSpan? initializeTimeout = null)
    {
        InfobasePath = infobasePath ?? throw new ArgumentNullException(nameof(infobasePath));
        Username = username ?? throw new ArgumentNullException(nameof(username));
        Password = password;
        InitializeTimeout = initializeTimeout;
    }

    public string InfobasePath { get; }

    public string Username { get; }

    public string? Password { get; }

    public TimeSpan? InitializeTimeout { get; }
}
