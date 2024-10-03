using System.Web;
using System.Web.Http;
using V77ApplicationWebApi.Web.Formatting;

namespace V77ApplicationWebApi.Web;

public class WebApiApplication : HttpApplication
{
    protected void Application_Start() => GlobalConfiguration.Configure(config =>
    {
        // Add "text/plain" media type support
        config.Formatters.Add(new PlainTextMediaTypeFormatter());

        config.MapHttpAttributeRoutes();
    });
}
