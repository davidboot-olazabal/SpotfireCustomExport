using System;
using System.Windows.Forms;
using SpotfireCustomExport.Core;

namespace SpotfireCustomExport.Forms
{
    public class ExportPromptForm : Form
    {
        private ExportPromptModel _model;
        private RadioButton _rbDataTable, _rbVisualization, _rbPage, _rbDocument;
        private ComboBox _itemDropdown, _formatDropdown;

        public ExportPromptForm(ExportPromptModel model)
        {
            _model = model;
            InitializeComponent();
            PopulateDropdowns();
        }

        private void PopulateDropdowns()
        {
            _itemDropdown.Items.Clear();

            if (_rbDataTable.Checked) { foreach (string item in _model.AvailableTables) _itemDropdown.Items.Add(item); }
            else if (_rbVisualization.Checked) { foreach (string item in _model.AvailableVisuals) _itemDropdown.Items.Add(item); }
            else if (_rbPage.Checked) { foreach (string item in _model.AvailablePages) _itemDropdown.Items.Add(item); }
            else if (_rbDocument.Checked) { foreach (string item in _model.AvailableDocuments) _itemDropdown.Items.Add(item); }

            if (_itemDropdown.Items.Count > 0) _itemDropdown.SelectedIndex = 0;
            UpdateFormats();
        }

        private void UpdateFormats()
        {
            string currentFormat = _formatDropdown.SelectedItem?.ToString();
            _formatDropdown.Items.Clear();

            if (_rbDataTable.Checked)
            {
                _formatDropdown.Items.AddRange(new string[] { "CSV", "TXT", "XLSX", "JSON", "XML" });
            }
            else if (_rbVisualization.Checked)
            {
                _formatDropdown.Items.Add("PNG");

                string selectedItem = _itemDropdown.SelectedItem?.ToString();
                bool isTabular = selectedItem != null && _model.TabularVisuals != null && _model.TabularVisuals.Contains(selectedItem);

                if (isTabular)
                {
                    _formatDropdown.Items.AddRange(new string[] { "XLSX", "JSON", "XML" });
                }
                else
                {
                    _formatDropdown.Items.Add("SVG");
                }
            }
            else if (_rbPage.Checked || _rbDocument.Checked)
            {
                _formatDropdown.Items.Add("PDF");
            }

            if (currentFormat != null && _formatDropdown.Items.Contains(currentFormat))
                _formatDropdown.SelectedItem = currentFormat;
            else if (_formatDropdown.Items.Count > 0)
                _formatDropdown.SelectedIndex = 0;
        }

        private void InitializeComponent()
        {
            this.Width = 400;
            this.Height = 280;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.Text = "Custom Export";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // SECTION 1: Source Selection
            Label sourceLabel = new Label() { Left = 20, Top = 15, Width = 340, Text = "1. Select Source Type:" };

            _rbDataTable = new RadioButton() { Left = 20, Top = 35, Width = 150, Text = "Data Table", Checked = _model.SourceType == "DataTable" };
            _rbVisualization = new RadioButton() { Left = 170, Top = 35, Width = 150, Text = "Visualization", Checked = _model.SourceType == "Visualization" };
            _rbPage = new RadioButton() { Left = 20, Top = 60, Width = 150, Text = "Page", Checked = _model.SourceType == "Page" };
            _rbDocument = new RadioButton() { Left = 170, Top = 60, Width = 150, Text = "Complete Report", Checked = _model.SourceType == "Document" };

            EventHandler updateEvt = (s, e) => { if (((RadioButton)s).Checked) PopulateDropdowns(); };
            _rbDataTable.CheckedChanged += updateEvt;
            _rbVisualization.CheckedChanged += updateEvt;
            _rbPage.CheckedChanged += updateEvt;
            _rbDocument.CheckedChanged += updateEvt;

            Label itemLabel = new Label() { Left = 20, Top = 90, Width = 340, Text = "2. Select Item to Export:" };
            _itemDropdown = new ComboBox() { Left = 20, Top = 110, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList };
            _itemDropdown.SelectedIndexChanged += (s, e) => UpdateFormats();

            Label formatLabel = new Label() { Left = 20, Top = 145, Width = 340, Text = "3. Select Export Format:" };
            _formatDropdown = new ComboBox() { Left = 20, Top = 165, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList };

            // BUTTONS
            Button confirmation = new Button() { Text = "Export", Left = 240, Width = 120, Top = 205 };
            Button cancel = new Button() { Text = "Cancel", Left = 110, Width = 120, Top = 205, DialogResult = DialogResult.Cancel };

            confirmation.Click += (sender, e) =>
            {
                _model.SourceType = _rbDataTable.Checked ? "DataTable" : _rbVisualization.Checked ? "Visualization" : _rbPage.Checked ? "Page" : "Document";
                _model.SelectedItem = _itemDropdown.SelectedItem?.ToString();
                _model.SelectedFormat = _formatDropdown.SelectedItem?.ToString();

                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.AddRange(new Control[] {
                sourceLabel, _rbDataTable, _rbVisualization, _rbPage, _rbDocument,
                itemLabel, _itemDropdown,
                formatLabel, _formatDropdown,
                confirmation, cancel
            });
            this.AcceptButton = confirmation;
            this.CancelButton = cancel;
        }
    }
}