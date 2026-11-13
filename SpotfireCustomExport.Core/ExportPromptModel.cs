using System.Collections.Generic;

namespace SpotfireCustomExport.Core
{
    public class ExportPromptModel
    {
        public string SourceType { get; set; } = "DataTable";

        public List<string> AvailableTables { get; set; } = new List<string>();
        public List<string> AvailableVisuals { get; set; } = new List<string>();
        public List<string> TabularVisuals { get; set; } = new List<string>();

        public List<string> AvailablePages { get; set; } = new List<string>();
        public List<string> AvailableDocuments { get; set; } = new List<string> { "Entire Analysis" };

        public string SelectedItem { get; set; }

        public List<string> AvailableFormats { get; set; } = new List<string>
        {
            "CSV", "TXT", "XLSX", "PDF", "PNG", "JSON", "XML", "SVG"
        };

        public string SelectedFormat { get; set; } = "CSV";
    }
}