using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Aspose_Assignment_App.Startup))]
namespace Aspose_Assignment_App
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
