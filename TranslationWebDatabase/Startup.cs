using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(TranslationWebDatabase.Startup))]
namespace TranslationWebDatabase
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
