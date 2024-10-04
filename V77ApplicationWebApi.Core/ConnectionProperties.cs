namespace V77ApplicationWebApi.Core;

public record ConnectionProperties
{
    public ConnectionProperties(string infobasePath, string username, string password)
    {
        InfobasePath = infobasePath;
        Username = username;
        Password = password;
    }

    public string InfobasePath { get; }

    public string Username { get; }

    public string Password { get; }
}
