using System.IO;
using Spotfire.Dxp.Web.Forms;
using SpotfireCustomExport.Core;
using System.Text;
using System.Linq;

namespace SpotfireCustomExport.Web
{
    public sealed class ExportPromptWebControl : CustomWizardPromptControl
    {
        public ExportPromptWebControl(ExportPromptModel model)
        {
            this.AddPage(new ExportSelectionPage(model));
        }

        private class ExportSelectionPage : CustomWizardPromptControlPage
        {
            private readonly ExportPromptModel _model;

            public ExportSelectionPage(ExportPromptModel model) : base("Direct Export")
            {
                _model = model;
            }

            protected override void OnGetContentsCore(TextWriter writer)
            {
                string tablesArray = string.Join(",", _model.AvailableTables.Select(t => $"'{t.Replace("'", "\\'")}'"));
                string visualsArray = string.Join(",", _model.AvailableVisuals.Select(v => $"'{v.Replace("'", "\\'")}'"));
                string tabularArray = _model.TabularVisuals != null ? string.Join(",", _model.TabularVisuals.Select(v => $"'{v.Replace("'", "\\'")}'")) : "";
                string pagesArray = string.Join(",", _model.AvailablePages.Select(p => $"'{p.Replace("'", "\\'")}'"));
                string docsArray = string.Join(",", _model.AvailableDocuments.Select(d => $"'{d.Replace("'", "\\'")}'"));

                writer.WriteLine($@"
                <style>
                    .export-group {{ margin-bottom: 6px; font-family: 'Segoe UI', sans-serif; font-size: 11px; }}
                    .export-label {{ font-weight: 600; display: block; margin-bottom: 2px; color: #222; }}
                    .export-control {{ width: 100%; box-sizing: border-box; padding: 2px 5px; font-family: 'Segoe UI', sans-serif; font-size: 11px; border: 1px solid #aaa; border-radius: 3px; height: 24px; }}
                    .wrap-container {{ width: 340px; margin: 0 auto; padding: 4px; overflow: hidden; }}
                    
                    .radio-group {{ display: flex; flex-wrap: wrap; margin-bottom: 6px; font-family: 'Segoe UI', sans-serif; font-size: 11px; }}
                    .radio-label {{ font-weight: 600; width: 100%; margin-bottom: 2px; color: #222; }}
                    .radio-option {{ width: 48%; margin-bottom: 3px; display: flex; align-items: center; cursor: pointer; }}
                    .radio-option input {{ margin: 0 4px 0 0; cursor: pointer; }}
                    
                    .flex-row {{ display: flex; justify-content: space-between; margin-bottom: 6px; }}
                    .col-main {{ width: 66%; }}
                    .col-sub {{ width: 31%; }}
                </style>
                
                <div class='wrap-container'>
                
                <input type='hidden' id='sourceType' name='sourceType' value='{_model.SourceType}' />

                <script>
                    var availableTables = [{tablesArray}];
                    var availableVisuals = [{visualsArray}];
                    var tabularVisuals = [{tabularArray}];
                    var availablePages = [{pagesArray}];
                    var availableDocs = [{docsArray}];

                    function getSourceType() {{
                        var radios = document.getElementsByName('sourceTypeRadio');
                        for(var i = 0; i < radios.length; i++) {{
                            if(radios[i].checked) return radios[i].value;
                        }}
                        return 'DataTable';
                    }}

                    function updateSourceType() {{
                        var type = getSourceType();
                        document.getElementById('sourceType').value = type;
                        
                        var itemSelect = document.getElementById('itemSelector');
                        itemSelect.innerHTML = '';
                        var items = [];
                        
                        if (type === 'DataTable') items = availableTables;
                        else if (type === 'Visualization') items = availableVisuals;
                        else if (type === 'Page') items = availablePages;
                        else if (type === 'Document') items = availableDocs;
                        
                        for (var i = 0; i < items.length; i++) {{
                            var opt = document.createElement('option');
                            opt.value = items[i];
                            opt.innerHTML = items[i];
                            itemSelect.appendChild(opt);
                        }}
                        updateFormats(); 
                    }}

                    function updateFormats() {{
                        var type = getSourceType();
                        var itemSelect = document.getElementById('itemSelector');
                        var formatSelect = document.getElementById('formatSelector');
                        
                        var currentFormat = formatSelect.value;
                        var currentItem = itemSelect.value;
                        
                        formatSelect.innerHTML = '';
                        var formats = [];
                        
                        if (type === 'DataTable') {{
                            formats = ['CSV', 'TXT', 'XLSX', 'JSON', 'XML'];
                        }} else if (type === 'Visualization') {{
                            formats = ['PNG'];
                            if (tabularVisuals.indexOf(currentItem) !== -1) {{
                                formats.push('XLSX');
                                formats.push('JSON');
                                formats.push('XML');
                            }} else {{
                                formats.push('SVG');
                            }}
                        }} else if (type === 'Page' || type === 'Document') {{
                            formats = ['PDF'];
                        }}
                        
                        for (var j = 0; j < formats.length; j++) {{
                            var opt = document.createElement('option');
                            opt.value = formats[j];
                            opt.innerHTML = formats[j];
                            formatSelect.appendChild(opt);
                        }}
                        
                        if (formats.indexOf(currentFormat) !== -1) {{
                            formatSelect.value = currentFormat;
                        }} else if (formats.length > 0) {{
                            formatSelect.value = formats[0];
                        }}
                    }}
                </script>
                
                <div class='radio-group'>
                    <span class='radio-label'>1. Source Type:</span>
                    <label class='radio-option'><input type='radio' name='sourceTypeRadio' value='DataTable' onchange='updateSourceType()' {(_model.SourceType == "DataTable" ? "checked" : "")}> Data Table</label>
                    <label class='radio-option'><input type='radio' name='sourceTypeRadio' value='Visualization' onchange='updateSourceType()' {(_model.SourceType == "Visualization" ? "checked" : "")}> Visualization</label>
                    <label class='radio-option'><input type='radio' name='sourceTypeRadio' value='Page' onchange='updateSourceType()' {(_model.SourceType == "Page" ? "checked" : "")}> Page</label>
                    <label class='radio-option'><input type='radio' name='sourceTypeRadio' value='Document' onchange='updateSourceType()' {(_model.SourceType == "Document" ? "checked" : "")}> Complete Report</label>
                </div>

                <div class='flex-row'>
                    <div class='col-main'>
                        <span class='export-label'>2. Item to Export:</span>
                        <select id='itemSelector' name='itemSelector' class='export-control' onchange='updateFormats()'></select>
                    </div>
                    <div class='col-sub'>
                        <span class='export-label'>3. Format:</span>
                        <select id='formatSelector' name='formatSelector' class='export-control'></select>
                    </div>
                </div>

                <script>updateSourceType();</script>
                </div>");
            }

            protected override bool OnValidatePromptCore(FormData data)
            {
                string sourceType = data.GetStringValue("sourceType");
                string format = data.GetStringValue("formatSelector");
                string item = data.GetStringValue("itemSelector");

                if (sourceType == "Visualization")
                {
                    bool isTabularFormat = format == "XLSX" || format == "JSON" || format == "XML";
                    bool isTabularItem = _model.TabularVisuals != null && _model.TabularVisuals.Contains(item);

                    if (isTabularFormat && !isTabularItem)
                    {
                        this.ErrorMessage = $"{format} export is only supported for tabular visualizations. '{item}' is not a table.";
                        return false;
                    }

                    if (format == "SVG" && isTabularItem)
                    {
                        this.ErrorMessage = $"SVG export is only supported for non-tabular visualizations. '{item}' is a table.";
                        return false;
                    }
                }

                return true;
            }

            protected override void OnLeavePageCore(FormData data)
            {
                _model.SourceType = data.GetStringValue("sourceType");
                _model.SelectedItem = data.GetStringValue("itemSelector");
                _model.SelectedFormat = data.GetStringValue("formatSelector");
            }

            protected override void OnGetScriptsCore(StringBuilder builder) { }
        }
    }
}