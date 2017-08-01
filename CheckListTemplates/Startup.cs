using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(CheckListTemplates.Startup))]
namespace CheckListTemplates
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
