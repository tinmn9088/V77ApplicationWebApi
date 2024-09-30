using System.Web;
using System.Web.Http;

namespace V77ApplicationWebApi.Web;

public class WebApiApplication : HttpApplication
{
    protected void Application_Start() => GlobalConfiguration.Configure(config => config.MapHttpAttributeRoutes());
}
