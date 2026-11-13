using Spotfire.Dxp.Application.Extension;
using Spotfire.Dxp.Web.Forms;
using SpotfireCustomExport.Core;

namespace SpotfireCustomExport.Web
{
    public class WebAddIn : AddIn
    {
        protected override void RegisterViews(ViewRegistrar registrar)
        {
            base.RegisterViews(registrar);
            registrar.Register(typeof(PromptControl), typeof(ExportPromptModel), typeof(ExportPromptWebControl));
        }
    }
}