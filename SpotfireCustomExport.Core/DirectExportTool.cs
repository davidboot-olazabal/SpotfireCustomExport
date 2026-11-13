using OfficeOpenXml;
using Spotfire.Dxp.Application;
using Spotfire.Dxp.Application.Export;
using Spotfire.Dxp.Application.Extension;
using Spotfire.Dxp.Application.Tools;
using Spotfire.Dxp.Application.Visuals;
using Spotfire.Dxp.Data;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace SpotfireCustomExport.Core
{
    public class DirectExportTool : CustomExportTool
    {
        public DirectExportTool() : base("Direct Export...", (ExtensionIcon)null, null) { }

        protected override IEnumerable<object> ExecuteCore(Document document, ExportResult result)
        {
            var model = new ExportPromptModel();

            // 1. Populate Available Tables
            foreach (DataTable table in document.Data.Tables)
            {
                model.AvailableTables.Add(table.Name);
            }

            // 2. Populate Available Visuals & Pages
            foreach (var page in document.Pages)
            {
                model.AvailablePages.Add(page.Title);

                foreach (var visual in page.Visuals)
                {
                    string vTitle = string.IsNullOrWhiteSpace(visual.Title) ? "[Empty Title]" : visual.Title.Trim();
                    model.AvailableVisuals.Add(vTitle);

                    if (visual.TypeId == VisualTypeIdentifiers.Table ||
                        visual.TypeId == VisualTypeIdentifiers.CrossTable ||
                        visual.TypeId == VisualTypeIdentifiers.SummaryTable)
                    {
                        model.TabularVisuals.Add(vTitle);
                    }
                }
            }

            if (model.AvailableTables.Count > 0) model.SelectedItem = model.AvailableTables[0];

            yield return model;

            // --- UPFRONT VALIDATIONS --- //

            if (model.SourceType == "Visualization")
            {
                bool isTabularFormat = model.SelectedFormat == "XLSX" || model.SelectedFormat == "JSON" || model.SelectedFormat == "XML";
                if (isTabularFormat)
                {
                    Visual validationVisual = FindVisualByTitle(document, model.SelectedItem);
                    if (validationVisual != null &&
                        validationVisual.TypeId != VisualTypeIdentifiers.Table &&
                        validationVisual.TypeId != VisualTypeIdentifiers.CrossTable &&
                        validationVisual.TypeId != VisualTypeIdentifiers.SummaryTable)
                    {
                        throw new InvalidOperationException($"Export Aborted: {model.SelectedFormat} export is only supported for tabular visualizations. '{model.SelectedItem}' is not a table. Please select PNG, PDF, or SVG instead.");
                    }
                }

                if (model.SelectedFormat == "SVG")
                {
                    Visual validationVisual = FindVisualByTitle(document, model.SelectedItem);
                    if (validationVisual != null &&
                        (validationVisual.TypeId == VisualTypeIdentifiers.Table ||
                         validationVisual.TypeId == VisualTypeIdentifiers.CrossTable ||
                         validationVisual.TypeId == VisualTypeIdentifiers.SummaryTable))
                    {
                        throw new InvalidOperationException($"Export Aborted: SVG export is only supported for non-tabular visualizations. '{model.SelectedItem}' is a table. Please select XLSX, CSV, JSON, or XML instead.");
                    }
                }
            }

            if (model.SourceType == "Page" || model.SourceType == "Document")
            {
                if (model.SelectedFormat != "PDF")
                {
                    throw new InvalidOperationException($"Export Aborted: Only PDF export is supported for full pages and complete reports.");
                }
            }

            string baseFileName = $"{model.SelectedItem.Replace(" ", "_")}_Export";

            // --- EXPORT EXECUTION --- //

            if (model.SourceType == "DataTable")
            {
                if (document.Data.Tables.TryGetValue(model.SelectedItem, out DataTable selectedTable))
                {
                    if (model.SelectedFormat == "TXT" || model.SelectedFormat == "CSV") ExportDelimitedText(selectedTable, result, baseFileName, model.SelectedFormat);
                    else if (model.SelectedFormat == "XLSX") ExportDataTableToExcel(selectedTable, result, baseFileName);
                    else if (model.SelectedFormat == "JSON") ExportDataTableToJson(selectedTable, result, baseFileName);
                    else if (model.SelectedFormat == "XML") ExportDataTableToXml(selectedTable, result, baseFileName);
                }
            }
            else if (model.SourceType == "Visualization")
            {
                Visual targetVisual = FindVisualByTitle(document, model.SelectedItem);
                if (targetVisual != null)
                {
                    if (model.SelectedFormat == "XLSX") ExportVisualToExcel(targetVisual, result, baseFileName);
                    else if (model.SelectedFormat == "PNG") ExportVisualToPng(targetVisual, result, baseFileName);
                    else if (model.SelectedFormat == "PDF") ExportGenericPdf(document, model, result, baseFileName);
                    else if (model.SelectedFormat == "JSON") ExportVisualToJson(targetVisual, result, baseFileName);
                    else if (model.SelectedFormat == "XML") ExportVisualToXml(targetVisual, result, baseFileName);
                    else if (model.SelectedFormat == "SVG") ExportVisualToSvg(targetVisual, result, baseFileName);
                }
            }
            else if (model.SourceType == "Page" || model.SourceType == "Document")
            {
                if (model.SelectedFormat == "PDF") ExportGenericPdf(document, model, result, baseFileName);
            }
        }

        // --- EXPORT METHODS --- //

        private void ExportDelimitedText(DataTable selectedTable, ExportResult result, string baseFileName, string format)
        {
            bool isTxt = format == "TXT";
            string extension = isTxt ? ".txt" : ".csv";
            string mimeType = isTxt ? "text/plain" : "text/csv";
            string delimiter = isTxt ? "\t" : ",";
            string fileName = baseFileName + extension;

            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType(mimeType));

            using (MemoryStream ms = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(ms, Encoding.UTF8, 1024, true))
                {
                    List<string> headers = new List<string>();
                    List<DataValueCursor<string>> cursors = new List<DataValueCursor<string>>();

                    foreach (DataColumn col in selectedTable.Columns)
                    {
                        headers.Add("\"" + col.Name.Replace("\"", "\"\"") + "\"");
                        cursors.Add(DataValueCursor.CreateFormatted(col));
                    }
                    writer.WriteLine(string.Join(delimiter, headers));

                    IndexSet allRows = new IndexSet(selectedTable.RowCount, true);
                    DataValueCursor[] cursorArray = cursors.ToArray();

                    foreach (DataRow row in selectedTable.GetRows(allRows, cursorArray))
                    {
                        List<string> rowValues = new List<string>();
                        foreach (var cursor in cursors)
                        {
                            string val = cursor.CurrentValue ?? "";
                            rowValues.Add("\"" + val.Replace("\"", "\"\"") + "\"");
                        }
                        writer.WriteLine(string.Join(delimiter, rowValues));
                    }
                    writer.Flush();
                }

                ms.Position = 0;
                using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
            }

            TriggerWinFormsFallback(result, dataPart, fileName, isTxt ? "Text files (*.txt)|*.txt|All files (*.*)|*.*" : "CSV files (*.csv)|*.csv|All files (*.*)|*.*");
        }

        private void ExportDataTableToExcel(DataTable selectedTable, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".xlsx";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

            OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("Spotfire Export Extension");

            using (MemoryStream ms = new MemoryStream())
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Exported Data");

                    int colIdx = 1;
                    List<DataValueCursor<string>> cursors = new List<DataValueCursor<string>>();
                    foreach (DataColumn col in selectedTable.Columns)
                    {
                        worksheet.Cells[1, colIdx].Value = col.Name;
                        cursors.Add(DataValueCursor.CreateFormatted(col));
                        colIdx++;
                    }

                    using (var range = worksheet.Cells[1, 1, 1, selectedTable.Columns.Count]) { range.Style.Font.Bold = true; }

                    IndexSet allRows = new IndexSet(selectedTable.RowCount, true);
                    DataValueCursor[] cursorArray = cursors.ToArray();
                    int rowIdx = 2;

                    foreach (DataRow row in selectedTable.GetRows(allRows, cursorArray))
                    {
                        for (int c = 0; c < cursors.Count; c++)
                        {
                            worksheet.Cells[rowIdx, c + 1].Value = cursors[c].CurrentValue ?? "";
                        }
                        rowIdx++;
                    }

                    worksheet.Cells.AutoFitColumns();
                    package.SaveAs(ms);
                }

                ms.Position = 0;
                using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*");
        }

        private void ExportDataTableToJson(DataTable selectedTable, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".json";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("application/json"));

            using (MemoryStream ms = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(ms, Encoding.UTF8, 1024, true))
                {
                    writer.WriteLine("[");
                    List<DataColumn> cols = selectedTable.Columns.Cast<DataColumn>().ToList();
                    List<DataValueCursor<string>> cursors = cols.Select(DataValueCursor.CreateFormatted).ToList();
                    IndexSet allRows = new IndexSet(selectedTable.RowCount, true);
                    DataValueCursor[] cursorArray = cursors.ToArray();

                    bool firstRow = true;
                    foreach (DataRow row in selectedTable.GetRows(allRows, cursorArray))
                    {
                        if (!firstRow) writer.WriteLine(",");
                        firstRow = false;

                        writer.Write("  {");
                        for (int c = 0; c < cols.Count; c++)
                        {
                            string colName = EscapeJsonString(cols[c].Name);
                            string val = cursors[c].CurrentValue ?? "";

                            writer.Write($"\"{colName}\": ");
                            if (TryParseNumeric(val, out double num))
                            {
                                writer.Write(num.ToString(System.Globalization.CultureInfo.InvariantCulture));
                            }
                            else if (bool.TryParse(val, out bool bVal))
                            {
                                writer.Write(bVal ? "true" : "false");
                            }
                            else
                            {
                                writer.Write($"\"{EscapeJsonString(val)}\"");
                            }

                            if (c < cols.Count - 1) writer.Write(", ");
                        }
                        writer.Write("}");
                    }
                    writer.WriteLine();
                    writer.WriteLine("]");
                    writer.Flush();
                }

                ms.Position = 0;
                using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "JSON files (*.json)|*.json|All files (*.*)|*.*");
        }

        private void ExportDataTableToXml(DataTable selectedTable, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".xml";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("application/xml"));

            using (MemoryStream ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, Indent = true, OmitXmlDeclaration = false };
                using (var xmlWriter = XmlWriter.Create(ms, settings))
                {
                    xmlWriter.WriteStartDocument();
                    xmlWriter.WriteStartElement("DataTable");
                    xmlWriter.WriteAttributeString("name", selectedTable.Name);

                    List<DataColumn> cols = selectedTable.Columns.Cast<DataColumn>().ToList();
                    List<DataValueCursor<string>> cursors = cols.Select(DataValueCursor.CreateFormatted).ToList();
                    IndexSet allRows = new IndexSet(selectedTable.RowCount, true);
                    DataValueCursor[] cursorArray = cursors.ToArray();

                    foreach (DataRow row in selectedTable.GetRows(allRows, cursorArray))
                    {
                        xmlWriter.WriteStartElement("Row");
                        for (int c = 0; c < cols.Count; c++)
                        {
                            string elemName = CleanXmlElementName(cols[c].Name);
                            string val = cursors[c].CurrentValue ?? "";
                            xmlWriter.WriteElementString(elemName, val);
                        }
                        xmlWriter.WriteEndElement();
                    }

                    xmlWriter.WriteEndElement();
                    xmlWriter.WriteEndDocument();
                    xmlWriter.Flush();
                }

                ms.Position = 0;
                using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "XML files (*.xml)|*.xml|All files (*.*)|*.*");
        }

        private void ExportVisualToExcel(Visual visual, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".xlsx";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

            OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("Spotfire Export Extension");

            using (MemoryStream ms = new MemoryStream())
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Exported Data");
                    string tempFile = Path.GetTempFileName();

                    try
                    {
                        if (visual.TypeId == VisualTypeIdentifiers.Table)
                        {
                            var tablePlot = visual.As<TablePlot>();
                            using (var writer = new StreamWriter(tempFile)) { tablePlot.ExportText(writer); }
                            LoadTextToWorksheet(tempFile, worksheet);

                            using (var range = worksheet.Cells[1, 1, 1, worksheet.Dimension?.End.Column ?? 1]) { range.Style.Font.Bold = true; }
                            ApplyDynamicTableColorings(tablePlot, worksheet);
                        }
                        else if (visual.TypeId == VisualTypeIdentifiers.CrossTable)
                        {
                            var crossTable = visual.As<CrossTablePlot>();
                            using (var writer = new StreamWriter(tempFile)) { crossTable.ExportText(writer); }
                            LoadTextToWorksheet(tempFile, worksheet);

                            if (worksheet.Dimension != null)
                            {
                                using (var range = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column]) { range.Style.Font.Bold = true; }
                                using (var range = worksheet.Cells[2, 1, worksheet.Dimension.End.Row, 1]) { range.Style.Font.Bold = true; }
                                worksheet.View.FreezePanes(2, 2);
                            }
                            ApplyDynamicCrossTableColorings(crossTable, worksheet);
                        }
                        else if (visual.TypeId == VisualTypeIdentifiers.SummaryTable)
                        {
                            var summaryTable = visual.As<SummaryTable>();
                            using (var writer = new StreamWriter(tempFile)) { summaryTable.ExportText(writer); }
                            LoadTextToWorksheet(tempFile, worksheet);

                            if (worksheet.Dimension != null)
                            {
                                using (var range = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column]) { range.Style.Font.Bold = true; }
                            }
                        }
                    }
                    finally
                    {
                        if (File.Exists(tempFile)) File.Delete(tempFile);
                    }

                    if (worksheet.Dimension != null) worksheet.Cells.AutoFitColumns();
                    package.SaveAs(ms);
                }

                ms.Position = 0;
                using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*");
        }

        private void ExportVisualToJson(Visual visual, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".json";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("application/json"));
            string tempFile = Path.GetTempFileName();

            try
            {
                ExportVisualToTextFile(visual, tempFile);
                if (!File.Exists(tempFile)) return;

                string[] lines = File.ReadAllLines(tempFile);
                if (lines.Length == 0) return;

                string[] headers = lines[0].Split('\t').Select(h => h.Trim('"').Trim()).ToArray();

                using (MemoryStream ms = new MemoryStream())
                {
                    using (StreamWriter writer = new StreamWriter(ms, Encoding.UTF8, 1024, true))
                    {
                        writer.WriteLine("[");
                        bool firstRow = true;

                        for (int r = 1; r < lines.Length; r++)
                        {
                            if (string.IsNullOrWhiteSpace(lines[r])) continue;
                            string[] cells = lines[r].Split('\t');

                            if (!firstRow) writer.WriteLine(",");
                            firstRow = false;

                            writer.Write("  {");
                            for (int c = 0; c < headers.Length; c++)
                            {
                                string hName = EscapeJsonString(headers[c]);
                                string val = c < cells.Length ? cells[c].Trim('"').Trim() : "";

                                writer.Write($"\"{hName}\": ");
                                if (TryParseNumeric(val, out double num))
                                {
                                    writer.Write(num.ToString(System.Globalization.CultureInfo.InvariantCulture));
                                }
                                else if (bool.TryParse(val, out bool bVal))
                                {
                                    writer.Write(bVal ? "true" : "false");
                                }
                                else
                                {
                                    writer.Write($"\"{EscapeJsonString(val)}\"");
                                }

                                if (c < headers.Length - 1) writer.Write(", ");
                            }
                            writer.Write("}");
                        }
                        writer.WriteLine();
                        writer.WriteLine("]");
                        writer.Flush();
                    }

                    ms.Position = 0;
                    using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
                }
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "JSON files (*.json)|*.json|All files (*.*)|*.*");
        }

        private void ExportVisualToXml(Visual visual, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".xml";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("application/xml"));
            string tempFile = Path.GetTempFileName();

            try
            {
                ExportVisualToTextFile(visual, tempFile);
                if (!File.Exists(tempFile)) return;

                string[] lines = File.ReadAllLines(tempFile);
                if (lines.Length == 0) return;

                string[] headers = lines[0].Split('\t').Select(h => h.Trim('"').Trim()).ToArray();

                using (MemoryStream ms = new MemoryStream())
                {
                    var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, Indent = true, OmitXmlDeclaration = false };
                    using (var xmlWriter = XmlWriter.Create(ms, settings))
                    {
                        xmlWriter.WriteStartDocument();
                        xmlWriter.WriteStartElement("VisualizationData");
                        xmlWriter.WriteAttributeString("title", visual.Title ?? "Visual");

                        for (int r = 1; r < lines.Length; r++)
                        {
                            if (string.IsNullOrWhiteSpace(lines[r])) continue;
                            string[] cells = lines[r].Split('\t');

                            xmlWriter.WriteStartElement("Row");
                            for (int c = 0; c < headers.Length; c++)
                            {
                                string elemName = CleanXmlElementName(headers[c]);
                                string val = c < cells.Length ? cells[c].Trim('"').Trim() : "";
                                xmlWriter.WriteElementString(elemName, val);
                            }
                            xmlWriter.WriteEndElement();
                        }

                        xmlWriter.WriteEndElement();
                        xmlWriter.WriteEndDocument();
                        xmlWriter.Flush();
                    }

                    ms.Position = 0;
                    using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
                }
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "XML files (*.xml)|*.xml|All files (*.*)|*.*");
        }

        private void ExportVisualToSvg(Visual visual, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".svg";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("image/svg+xml"));

            int width = 2400;
            int height = 1600;

            using (Bitmap bitmap = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
#pragma warning disable CS0618
                    visual.Render(g, new Rectangle(0, 0, width, height));
#pragma warning restore CS0618
                }

                using (MemoryStream pngStream = new MemoryStream())
                {
                    bitmap.Save(pngStream, ImageFormat.Png);
                    string base64 = Convert.ToBase64String(pngStream.ToArray());

                    using (MemoryStream svgStream = new MemoryStream())
                    {
                        using (StreamWriter sw = new StreamWriter(svgStream, Encoding.UTF8, 1024, true))
                        {
                            sw.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                            sw.WriteLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" viewBox=\"0 0 {width} {height}\" width=\"100%\" height=\"100%\">");
                            sw.WriteLine($"  <title>{SecurityElement.Escape(visual.Title ?? "Visualization")}</title>");
                            sw.WriteLine($"  <image width=\"{width}\" height=\"{height}\" xlink:href=\"data:image/png;base64,{base64}\"/>");
                            sw.WriteLine("</svg>");
                            sw.Flush();
                        }

                        svgStream.Position = 0;
                        using (var partStream = dataPart.GetStream()) { svgStream.CopyTo(partStream); }
                    }
                }
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "SVG Vector Images (*.svg)|*.svg|All files (*.*)|*.*");
        }

        private void ExportVisualToPng(Visual visual, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".png";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("image/png"));

            using (MemoryStream ms = new MemoryStream())
            {
                using (Bitmap bitmap = new Bitmap(1200, 800))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
#pragma warning disable CS0618
                        visual.Render(g, new Rectangle(0, 0, 1200, 800));
#pragma warning restore CS0618
                    }
                    bitmap.Save(ms, ImageFormat.Png);
                }

                ms.Position = 0;
                using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "PNG Image (*.png)|*.png|All files (*.*)|*.*");
        }

        private void ExportGenericPdf(Document document, ExportPromptModel model, ExportResult result, string baseFileName)
        {
            string fileName = baseFileName + ".pdf";
            ExportResultPart dataPart = result.AddPart(fileName, new System.Net.Mime.ContentType("application/pdf"));
            string tempFile = Path.GetTempFileName() + ".pdf";

            try
            {
#pragma warning disable CS0618
                PdfExportSettings pdfSettings = new PdfExportSettings();

                if (model.SourceType == "Visualization")
                {
                    Visual targetVisual = FindVisualByTitle(document, model.SelectedItem);
                    if (targetVisual != null)
                    {
                        document.ActivePageReference = targetVisual.Context.GetAncestor<Page>();
                        pdfSettings.Scope = ExportScope.ActiveVisualization;
                    }
                }
                else if (model.SourceType == "Page")
                {
                    Page targetPage = FindPageByTitle(document, model.SelectedItem);
                    if (targetPage != null)
                    {
                        document.ActivePageReference = targetPage;
                        pdfSettings.Scope = ExportScope.ActivePage;
                    }
                }
                else if (model.SourceType == "Document")
                {
                    pdfSettings.Scope = ExportScope.AllPages;
                }

                document.Export(pdfSettings, tempFile);
#pragma warning restore CS0618

                if (File.Exists(tempFile))
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (FileStream fs = File.OpenRead(tempFile))
                        {
                            fs.CopyTo(ms);
                        }
                        ms.Position = 0;
                        using (var partStream = dataPart.GetStream()) { ms.CopyTo(partStream); }
                    }
                }
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }

            TriggerWinFormsFallback(result, dataPart, fileName, "PDF Document (*.pdf)|*.pdf|All files (*.*)|*.*");
        }

        // --- HELPER METHODS --- //

        private void ExportVisualToTextFile(Visual visual, string tempFile)
        {
            using (var writer = new StreamWriter(tempFile))
            {
                if (visual.TypeId == VisualTypeIdentifiers.Table) visual.As<TablePlot>().ExportText(writer);
                else if (visual.TypeId == VisualTypeIdentifiers.CrossTable) visual.As<CrossTablePlot>().ExportText(writer);
                else if (visual.TypeId == VisualTypeIdentifiers.SummaryTable) visual.As<SummaryTable>().ExportText(writer);
            }
        }

        private string EscapeJsonString(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        }

        private string CleanXmlElementName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Field";
            string clean = Regex.Replace(name.Trim(), @"[^\w.-]", "_");
            if (char.IsDigit(clean[0])) clean = "_" + clean;
            return clean;
        }

        private Visual FindVisualByTitle(Document document, string title)
        {
            foreach (var page in document.Pages)
            {
                foreach (var visual in page.Visuals)
                {
                    string vTitle = string.IsNullOrWhiteSpace(visual.Title) ? "[Empty Title]" : visual.Title.Trim();
                    if (vTitle.Equals(title, StringComparison.OrdinalIgnoreCase)) return visual;
                }
            }
            return null;
        }

        private Page FindPageByTitle(Document document, string title)
        {
            foreach (var page in document.Pages)
            {
                if (page.Title.Equals(title, StringComparison.OrdinalIgnoreCase)) return page;
            }
            return null;
        }

        private bool TryParseNumeric(string val, out double result)
        {
            result = 0;
            if (string.IsNullOrWhiteSpace(val)) return false;

            string clean = new string(val.Where(c => char.IsDigit(c) || c == ',' || c == '.' || c == '-').ToArray());
            if (clean.Length == 0 || clean == "-") return false;

            int lastComma = clean.LastIndexOf(',');
            int lastDot = clean.LastIndexOf('.');

            if (lastComma == -1 && lastDot == -1)
            {
                return double.TryParse(clean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out result);
            }

            char decimalSeparator = (lastComma > lastDot) ? ',' : '.';
            string digitsOnly = "";
            foreach (char c in clean)
            {
                if (c == decimalSeparator) digitsOnly += '.';
                else if (char.IsDigit(c) || c == '-') digitsOnly += c;
            }

            return double.TryParse(digitsOnly, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out result);
        }

        private void LoadTextToWorksheet(string filePath, ExcelWorksheet worksheet)
        {
            if (!File.Exists(filePath)) return;
            string[] lines = File.ReadAllLines(filePath);
            for (int r = 0; r < lines.Length; r++)
            {
                string[] cells = lines[r].Split('\t');
                for (int c = 0; c < cells.Length; c++)
                {
                    string val = cells[c].Trim('"').Trim();
                    if (string.IsNullOrWhiteSpace(val)) continue;

                    if (TryParseNumeric(val, out double num))
                    {
                        worksheet.Cells[r + 1, c + 1].Value = num;
                        if (val.Contains("€")) worksheet.Cells[r + 1, c + 1].Style.Numberformat.Format = "€ #,##0.00";
                        else if (val.Contains("$")) worksheet.Cells[r + 1, c + 1].Style.Numberformat.Format = "$ #,##0.00";
                        else if (val.Contains("%")) worksheet.Cells[r + 1, c + 1].Style.Numberformat.Format = "0.00%";
                        else if (val.Contains(",") || val.Contains(".")) worksheet.Cells[r + 1, c + 1].Style.Numberformat.Format = "#,##0.00";
                    }
                    else
                    {
                        worksheet.Cells[r + 1, c + 1].Value = val;
                    }
                }
            }
        }

        private void TriggerWinFormsFallback(ExportResult result, ExportResultPart dataPart, string fileName, string filter)
        {
            if (!result.CanFrameworkHandleResult)
            {
                Type saveFileDialogType = Type.GetType("System.Windows.Forms.SaveFileDialog, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
                if (saveFileDialogType != null)
                {
                    dynamic sfd = Activator.CreateInstance(saveFileDialogType);
                    sfd.Filter = filter;
                    sfd.FileName = fileName;

                    if (sfd.ShowDialog()?.ToString() == "OK")
                    {
                        using (var fileStream = File.Create(sfd.FileName))
                        using (var partStream = dataPart.GetStream())
                        {
                            partStream.Position = 0;
                            partStream.CopyTo(fileStream);
                        }
                    }
                }
            }
        }

        // --- COLOR EXTRACTION METHODS --- //

        private string ForceReadProperty(object obj, params string[] names)
        {
            if (obj == null) return string.Empty;
            Type type = obj.GetType();
            var flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.FlattenHierarchy;

            Type currentType = type;
            while (currentType != null && currentType != typeof(object))
            {
                foreach (var name in names)
                {
                    try
                    {
                        var prop = currentType.GetProperty(name, flags);
                        if (prop != null) return prop.GetValue(obj)?.ToString() ?? "";
                        var field = currentType.GetField(name, flags);
                        if (field != null) return field.GetValue(obj)?.ToString() ?? "";
                    }
                    catch { }
                }
                currentType = currentType.BaseType;
            }

            foreach (var iface in type.GetInterfaces())
            {
                foreach (var name in names)
                {
                    try
                    {
                        var prop = iface.GetProperty(name, flags);
                        if (prop != null) return prop.GetValue(obj)?.ToString() ?? "";
                    }
                    catch { }
                }
            }
            return string.Empty;
        }

        private bool CheckIfSegments(object obj)
        {
            if (obj == null) return false;
            try
            {
                var flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.FlattenHierarchy;
                foreach (var prop in obj.GetType().GetProperties(flags))
                {
                    try
                    {
                        string val = prop.GetValue(obj)?.ToString();
                        if (!string.IsNullOrEmpty(val) && val.IndexOf("Segments", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                    }
                    catch { }
                }
                foreach (var field in obj.GetType().GetFields(flags))
                {
                    try
                    {
                        string val = field.GetValue(obj)?.ToString();
                        if (!string.IsNullOrEmpty(val) && val.IndexOf("Segments", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                    }
                    catch { }
                }
            }
            catch { }
            return false;
        }

        private double ParseSpotfireNumber(string valStr, double fallback)
        {
            if (string.IsNullOrWhiteSpace(valStr)) return fallback;
            if (valStr.IndexOf("Min", StringComparison.OrdinalIgnoreCase) >= 0) return -9999999999;
            if (valStr.IndexOf("Max", StringComparison.OrdinalIgnoreCase) >= 0) return 9999999999;
            if (valStr.Contains("Spotfire.Dxp")) return fallback;

            if (TryParseNumeric(valStr, out double parsed)) return parsed;

            return fallback;
        }

        private double ExtractNumericValue(object condObj)
        {
            if (condObj == null) return 0;

            string isMinStr = ForceReadProperty(condObj, "IsMin");
            if (isMinStr == "True") return -9999999999;

            string isMaxStr = ForceReadProperty(condObj, "IsMax");
            if (isMaxStr == "True") return 9999999999;

            string[] propsToCheck = new string[] { "Value", "AbsoluteValue", "Expression", "FormattedValue", "NumericValue" };

            foreach (string propName in propsToCheck)
            {
                string rawVal = ForceReadProperty(condObj, propName);
                if (!string.IsNullOrWhiteSpace(rawVal) && !rawVal.Contains("Spotfire.Dxp"))
                {
                    double parsed = ParseSpotfireNumber(rawVal, -12345.67);
                    if (parsed != -12345.67) return parsed;
                }
            }

            Type t = condObj.GetType();
            var flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.FlattenHierarchy;
            while (t != null && t != typeof(object))
            {
                foreach (var field in t.GetFields(flags))
                {
                    try
                    {
                        object fVal = field.GetValue(condObj);
                        if (fVal != null)
                        {
                            string fStr = fVal.ToString();
                            if (!string.IsNullOrWhiteSpace(fStr) && !fStr.Contains("Spotfire.Dxp"))
                            {
                                double parsed = ParseSpotfireNumber(fStr, -12345.67);
                                if (parsed != -12345.67) return parsed;
                            }
                        }
                    }
                    catch { }
                }
                t = t.BaseType;
            }

            return 0;
        }

        private double GetLuminance(Color c) { return (0.2126 * c.R) + (0.7152 * c.G) + (0.0722 * c.B); }

        private void ApplyDynamicTableColorings(TablePlot table, ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null) return;

            dynamic colorings = table.Colorings;
            int endRow = worksheet.Dimension.End.Row;
            int endCol = worksheet.Dimension.End.Column;

            foreach (dynamic coloring in colorings)
            {
                try
                {
                    string expr = ForceReadProperty(coloring, "Expression").Replace("[", "").Replace("]", "").Trim();
                    string colName = ForceReadProperty(coloring, "Name");
                    string dispName = ForceReadProperty(coloring, "DisplayName");

                    string target = ForceReadProperty(coloring, "ColorTarget", "ApplyColorTo", "CellColorTarget");
                    bool isText = target.IndexOf("Text", StringComparison.OrdinalIgnoreCase) >= 0;

                    int targetCol = -1;
                    for (int c = 1; c <= endCol; c++)
                    {
                        string headerName = worksheet.Cells[1, c].Text?.Trim();
                        if (!string.IsNullOrEmpty(headerName) &&
                            (string.Equals(headerName, expr, StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(headerName, colName, StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(headerName, dispName, StringComparison.OrdinalIgnoreCase)))
                        {
                            targetCol = c;
                            break;
                        }
                    }

                    if (targetCol != -1)
                    {
                        ApplyDynamicRules(worksheet, coloring, 2, targetCol, endRow, targetCol, true, isText, false);
                    }
                }
                catch { }
            }
        }

        private void ApplyDynamicCrossTableColorings(CrossTablePlot crossTable, ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null) return;
            try
            {
                dynamic colorings = crossTable.Colorings;
                if (colorings.Count > 0)
                {
                    dynamic mainColoring = colorings[0];
                    bool isText = false;
                    bool isSegments = false;

                    object colorAxis = null;
                    try { colorAxis = crossTable.GetType().GetProperty("ColorAxis", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.FlattenHierarchy)?.GetValue(crossTable); } catch { }

                    if (colorAxis != null)
                    {
                        object ctColoring = null;
                        try { ctColoring = colorAxis.GetType().GetProperty("Coloring", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.FlattenHierarchy)?.GetValue(colorAxis); } catch { }

                        if (ctColoring != null)
                        {
                            string target = ForceReadProperty(ctColoring, "ColorTarget", "ApplyColorTo", "CellColorTarget");
                            if (target.IndexOf("Text", StringComparison.OrdinalIgnoreCase) >= 0) isText = true;
                        }

                        if (CheckIfSegments(colorAxis)) isSegments = true;
                    }

                    int endRow = worksheet.Dimension.End.Row;
                    int endCol = worksheet.Dimension.End.Column;

                    int dataStartCol = 1;
                    for (int c = 1; c <= endCol; c++)
                    {
                        if (worksheet.Cells[endRow, c].Value is double)
                        {
                            dataStartCol = c;
                            break;
                        }
                    }
                    if (dataStartCol == 1) dataStartCol = 2;

                    if (endRow >= 2 && endCol >= dataStartCol)
                    {
                        ApplyDynamicRules(worksheet, mainColoring, 2, dataStartCol, endRow, endCol, true, isText, isSegments);
                    }
                }
            }
            catch { }
        }

        private void ApplyDynamicRules(ExcelWorksheet worksheet, dynamic coloring, int startRow, int startCol, int endRow, int endCol, bool allowDefaultColor, bool isText, bool isSegmentsPassed)
        {
            var dataAddress = new ExcelAddress(startRow, startCol, endRow, endCol);
            bool isSegments = isSegmentsPassed;

            if (allowDefaultColor)
            {
                try
                {
                    dynamic sfDefault = coloring.DefaultColor;
                    Color defaultColor = Color.FromArgb(255, (int)sfDefault.R, (int)sfDefault.G, (int)sfDefault.B);
                    bool isWhite = defaultColor.R == 255 && defaultColor.G == 255 && defaultColor.B == 255;

                    if (!isWhite)
                    {
                        if (!isText && GetLuminance(defaultColor) < 110)
                        {
                            isText = true;
                        }

                        using (var range = worksheet.Cells[startRow, startCol, endRow, endCol])
                        {
                            if (isText) { range.Style.Font.Color.SetColor(defaultColor); }
                            else
                            {
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(defaultColor);
                            }
                        }
                    }
                }
                catch { }
            }

            foreach (dynamic rule in coloring)
            {
                string typeName = rule.GetType().Name;
                if (CheckIfSegments(rule)) isSegments = true;

                if (typeName.IndexOf("Continuous", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    typeName.IndexOf("Range", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try
                    {
                        var pts = rule.Breakpoints;
                        var bpList = new List<Tuple<double, Color>>();

                        foreach (dynamic pt in pts)
                        {
                            double vNum = ExtractNumericValue(pt.Value);
                            Color ptColor = Color.FromArgb(255, (int)pt.Color.R, (int)pt.Color.G, (int)pt.Color.B);
                            bpList.Add(new Tuple<double, Color>(vNum, ptColor));
                        }

                        bpList = bpList.OrderBy(b => b.Item1).ToList();

                        if (isSegments)
                        {
                            try
                            {
                                var blankRule = worksheet.ConditionalFormatting.AddExpression(dataAddress);
                                blankRule.Formula = $"ISBLANK({dataAddress.Start.Address})";
                                blankRule.StopIfTrue = true;
                            }
                            catch { }

                            if (bpList.Count > 0)
                            {
                                bpList[0] = new Tuple<double, Color>(-9999999999, bpList[0].Item2);
                                bpList[bpList.Count - 1] = new Tuple<double, Color>(9999999999, bpList[bpList.Count - 1].Item2);
                            }

                            for (int i = 0; i < bpList.Count - 1; i++)
                            {
                                double v1 = bpList[i].Item1;
                                double v2 = bpList[i + 1].Item1;
                                Color segColor = bpList[i].Item2;

                                var ruleCond = worksheet.ConditionalFormatting.AddBetween(dataAddress);
                                ruleCond.Formula = v1.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                ruleCond.Formula2 = v2.ToString(System.Globalization.CultureInfo.InvariantCulture);

                                if (isText) { ruleCond.Style.Font.Color.Color = segColor; }
                                else
                                {
                                    ruleCond.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    ruleCond.Style.Fill.BackgroundColor.Color = segColor;
                                }
                            }
                        }
                        else
                        {
                            if (bpList.Count == 2)
                            {
                                var colorScale = worksheet.ConditionalFormatting.AddTwoColorScale(dataAddress);
                                colorScale.LowValue.Type = OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType.Min;
                                colorScale.LowValue.Color = bpList[0].Item2;
                                colorScale.HighValue.Type = OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType.Max;
                                colorScale.HighValue.Color = bpList[1].Item2;
                            }
                            else if (bpList.Count >= 3)
                            {
                                var colorScale = worksheet.ConditionalFormatting.AddThreeColorScale(dataAddress);
                                colorScale.LowValue.Type = OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType.Min;
                                colorScale.LowValue.Color = bpList[0].Item2;
                                colorScale.MiddleValue.Type = OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType.Percent;
                                colorScale.MiddleValue.Value = 50;
                                colorScale.MiddleValue.Color = bpList[bpList.Count / 2].Item2;
                                colorScale.HighValue.Type = OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType.Max;
                                colorScale.HighValue.Color = bpList[bpList.Count - 1].Item2;
                            }
                        }
                    }
                    catch { }
                    continue;
                }

                if (typeName.IndexOf("SingleColor", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try
                    {
                        dynamic sfColor = rule.Color;
                        using (var range = worksheet.Cells[startRow, startCol, endRow, endCol])
                        {
                            if (isText) { range.Style.Font.Color.SetColor(Color.FromArgb(255, (int)sfColor.R, (int)sfColor.G, (int)sfColor.B)); }
                            else
                            {
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, (int)sfColor.R, (int)sfColor.G, (int)sfColor.B));
                            }
                        }
                    }
                    catch { }
                    continue;
                }

                bool extractedAny = false;
                try
                {
                    var categories = rule.GetExplicitCategories();
                    foreach (var cat in categories)
                    {
                        dynamic sfColor = null;
                        try { sfColor = rule[cat]; } catch { }
                        if (sfColor != null)
                        {
                            ApplyStaticColorToGrid(worksheet, startRow, startCol, endRow, endCol, cat, sfColor, isText);
                            extractedAny = true;
                        }
                    }
                }
                catch { }

                if (!extractedAny)
                {
                    try
                    {
                        var keys = rule.GetCategoryKeys();
                        foreach (var key in keys)
                        {
                            dynamic sfColor = null;
                            try { sfColor = rule[key]; } catch { }
                            if (sfColor != null)
                            {
                                ApplyStaticColorToGrid(worksheet, startRow, startCol, endRow, endCol, key, sfColor, isText);
                            }
                        }
                    }
                    catch { }
                }
            }
        }

        private void ApplyStaticColorToGrid(ExcelWorksheet worksheet, int startRow, int startCol, int endRow, int endCol, dynamic sfCategory, dynamic sfColor, bool isText)
        {
            if (sfColor == null || sfCategory == null) return;

            string targetValue = sfCategory.ToString().Trim();
            if (targetValue.StartsWith("[") && targetValue.EndsWith("]"))
                targetValue = targetValue.Substring(1, targetValue.Length - 2).Trim();

            if (string.IsNullOrEmpty(targetValue)) return;

            Color color = Color.FromArgb(255, (int)sfColor.R, (int)sfColor.G, (int)sfColor.B);

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    string cellText = cell.Value?.ToString()?.Trim();

                    if (string.Equals(cellText, targetValue, StringComparison.OrdinalIgnoreCase))
                    {
                        if (isText) { cell.Style.Font.Color.SetColor(color); }
                        else
                        {
                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(color);
                        }
                    }
                }
            }
        }
    }
}