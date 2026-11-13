using System.Windows.Forms;
using Spotfire.Dxp.Application.Extension;
using SpotfireCustomExport.Core;

namespace SpotfireCustomExport.Forms
{
    public class FormsAddIn : AddIn
    {
        protected override void RegisterViews(ViewRegistrar registrar)
        {
            base.RegisterViews(registrar);
            registrar.Register(typeof(Form), typeof(ExportPromptModel), typeof(ExportPromptForm));
        }
    }
}