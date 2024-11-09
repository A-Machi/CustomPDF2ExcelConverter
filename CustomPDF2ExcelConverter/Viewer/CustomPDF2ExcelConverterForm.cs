using CustomPDF2ExcelConverter.Controller;
using CustomPDF2ExcelConverter.Model;
using CustomPDF2ExcelConverter.Viewer;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace CustomPDF2ExcelConverter
{
    public partial class CustomPDF2ExcelConverterForm : Form
    {

        private TextBox txtFilePathPdf = new TextBox();
        private TextBox txtFilePathExcel = new TextBox();
        private Button btnBrowsePdf = new Button();
        private Button btnBrowseExcel = new Button();
        private Button btnMerge = new Button();
        private Label lblFilePathPdf = new Label();
        private Label lblFilePathExcel = new Label();

        public CustomPDF2ExcelConverterForm()
        {
            InitializeComponent();
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            // Set the size of the text fields and buttons so they fit well side by side
            var textBoxWidth = 220;
            var buttonWidth = 120;
            var controlHeight = 25;
            var vSpacing = 20;
            var labelHeightPDF = 15; 
            var labelHeightExcel = 75; 

            // Initialize TextBox for PDF
            this.txtFilePathPdf = new TextBox();
            this.txtFilePathPdf.Location = new Point(20, labelHeightPDF + 20);
            this.txtFilePathPdf.Name = "txtFilePathPdf";
            this.txtFilePathPdf.Size = new Size(textBoxWidth, controlHeight);
            this.txtFilePathPdf.TabIndex = 0;
            this.Controls.Add(this.txtFilePathPdf);

            // Initialize Button for PDF browsing
            this.btnBrowsePdf = new Button();
            this.btnBrowsePdf.Location = new Point(250, labelHeightPDF + 20);
            this.btnBrowsePdf.Name = "btnBrowsePdf";
            this.btnBrowsePdf.Size = new Size(buttonWidth, controlHeight);
            this.btnBrowsePdf.TabIndex = 1;
            this.btnBrowsePdf.Text = "Browse To PDF";
            this.btnBrowsePdf.UseVisualStyleBackColor = true;
            this.btnBrowsePdf.Click += new EventHandler(this.btnBrowsePdf_Click!);
            this.Controls.Add(this.btnBrowsePdf);

            // Initialize Label for the TextBox PDF
            this.lblFilePathPdf = new Label();
            this.lblFilePathPdf.Location = new Point(18, labelHeightPDF);
            this.lblFilePathPdf.Name = "lblFilePathPdf";
            this.lblFilePathPdf.Size = new Size(200, controlHeight);
            this.lblFilePathPdf.Text = "Select PDF-file to convert:";
            this.Controls.Add(this.lblFilePathPdf);

            // Initialize TextBox for Excel
            this.txtFilePathExcel = new TextBox();
            this.txtFilePathExcel.Location = new Point(20, labelHeightExcel + 20);
            this.txtFilePathExcel.Name = "txtFilePathExcel";
            this.txtFilePathExcel.Size = new Size(textBoxWidth, controlHeight);
            this.txtFilePathExcel.TabIndex = 2;
            this.Controls.Add(this.txtFilePathExcel);

            // Initialize Button for Excel browsing
            this.btnBrowseExcel = new Button();
            this.btnBrowseExcel.Location = new Point(250, labelHeightExcel + 20);
            this.btnBrowseExcel.Name = "btnBrowseExcel";
            this.btnBrowseExcel.Size = new Size(buttonWidth, controlHeight);
            this.btnBrowseExcel.TabIndex = 3;
            this.btnBrowseExcel.Text = "Browse To Excel";
            this.btnBrowseExcel.UseVisualStyleBackColor = true;
            this.btnBrowseExcel.Click += new EventHandler(this.btnBrowseExcel_Click!);
            this.Controls.Add(this.btnBrowseExcel);

            // Initialize Label for the TextBox Excel
            this.lblFilePathExcel = new Label();
            this.lblFilePathExcel.Location = new Point(18, labelHeightExcel);
            this.lblFilePathExcel.Name = "lblFilePathExcel";
            this.lblFilePathExcel.Size = new Size(200, controlHeight);
            this.lblFilePathExcel.Text = "Select Excel-file to store conversion:";
            this.Controls.Add(this.lblFilePathExcel);

            // Initialize Button for Convert
            this.btnMerge = new Button();
            this.btnMerge.Location = new Point(250, 130 + vSpacing);
            this.btnMerge.Name = "btnMerge";
            this.btnMerge.Size = new Size(buttonWidth, 30);
            this.btnMerge.TabIndex = 4;
            this.btnMerge.Text = "Convert";
            this.btnMerge.UseVisualStyleBackColor = true;
            this.btnMerge.BackColor = ColorTranslator.FromHtml("#FDAE44");
            this.btnMerge.Click += new EventHandler(this.btnConvert_Click!);
            this.Controls.Add(this.btnMerge);
        }

        private void btnBrowsePdf_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
                openFileDialog.Title = "Browse PDF Files";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePath = openFileDialog.FileName;

                    txtFilePathPdf.Text = filePath;
                }
            }
        }

        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
                openFileDialog.Title = "Browse Excel Files";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePath = openFileDialog.FileName;

                    txtFilePathExcel.Text = filePath;
                }
            }
        }

        private async void btnConvert_Click(object sender, EventArgs e)
        {
            var pdfFilePath = txtFilePathPdf.Text;
            var excelFilePath = txtFilePathExcel.Text;

            if (string.IsNullOrEmpty(pdfFilePath) || string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Please select both PDF and Excel files before merging.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var loadingForm = new LoadingForm();

            try
            {
                loadingForm.Show();

                var errorMessage = string.Empty;

                var success = await Task.Run(() => {
                    try
                    {
                        return ConvertPDF2Excel.CustomPDF2ExcelConverterHandler(pdfFilePath, excelFilePath);
                    }
                    catch (Exception ex)
                    {
                        errorMessage = ex.Message;
                        return false;
                    }
                });

                if (success)
                {
                    MessageBox.Show("Convert operation completed successfully.", "Convert", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ConvertPDF2Excel.StartExcel(excelFilePath);
                }
                else
                {
                    MessageBox.Show($"Convert operation failed. Reason: {errorMessage}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                loadingForm.Hide();
                loadingForm.Dispose();
            }
        }
    }
}
