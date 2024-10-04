using System;

namespace V77ApplicationWebApi.Core;

public record ConnectionProperties
{
    public ConnectionProperties(string infobasePath, string username, string password)
    {
        InfobasePath = infobasePath ?? throw new ArgumentNullException(nameof(infobasePath));
        Username = username ?? throw new ArgumentNullException(nameof(username));
        Password = password ?? throw new ArgumentNullException(nameof(password));
    }

    public string InfobasePath { get; }

    public string Username { get; }

    public string Password { get; }
}
