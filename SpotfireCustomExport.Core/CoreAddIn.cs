using Spotfire.Dxp.Application.Extension;

namespace SpotfireCustomExport.Core
{
    public class CoreAddIn : AddIn
    {
        protected override void RegisterTools(ToolRegistrar registrar)
        {
            base.RegisterTools(registrar);
            registrar.Register(new DirectExportTool());
        }
    }
}