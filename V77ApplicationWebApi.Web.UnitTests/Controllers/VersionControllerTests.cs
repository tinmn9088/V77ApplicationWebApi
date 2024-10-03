using Xunit;

namespace V77ApplicationWebApi.Web.Controllers.UnitTests;

public class VersionControllerTests
{
    private readonly VersionController _controller;

    public VersionControllerTests() => _controller = new();

    [Fact]
    public void VersionController_GetAssemblyVersion_ShouldReturnNotNul()
    {
        // Act
        string assemblyVersion = _controller.GetAssemblyVersion();

        // Assert
        Assert.NotNull(assemblyVersion);
    }
}
