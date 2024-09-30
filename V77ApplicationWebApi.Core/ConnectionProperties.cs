namespace V77ApplicationWebApi.Core;

public class ConnectionProperties(string infobasePath, string username, string password)
{
    public string InfobasePath { get; } = infobasePath;

    public string Username { get; } = username;

    public string Password { get; } = password;
}
