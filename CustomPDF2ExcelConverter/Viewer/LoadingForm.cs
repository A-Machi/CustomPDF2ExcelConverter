namespace CustomPDF2ExcelConverter.Viewer
{
    public partial class LoadingForm : Form
    {
        public LoadingForm()
        {
            //InitializeComponent();
            SetupLoadingForm();
        }

        private void SetupLoadingForm()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.Size = new System.Drawing.Size(180, 50);

            Label lblLoading = new Label();
            lblLoading.Text = "Loading, please wait...";
            lblLoading.AutoSize = true;
            lblLoading.Location = new System.Drawing.Point(50, 15);

            this.Controls.Add(lblLoading);
        }
    }
}
