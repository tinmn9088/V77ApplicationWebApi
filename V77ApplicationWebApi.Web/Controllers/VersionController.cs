using System.Reflection;
using System.Web.Http;

namespace V77ApplicationWebApi.Web.Controllers;

[RoutePrefix("api/version")]
public class VersionController : ApiController
{
    private static readonly string s_assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();

    [HttpGet]
    [Route("")]
    public string GetAssemblyVersion() => s_assemblyVersion;
}
