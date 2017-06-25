using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace Dataentry
{
    public partial class MainWindow : Form
    {
        Excel.Application xlApp;
        BackgroundWork backgroundWork;
        public MainWindow()
        {
            
           InitializeComponent();
            SkippedItemsList.View = View.List;

        }

        private void ConvertToExcelButton_Click(object sender, EventArgs e)
        {
            String fileToConvert = TextFilePathtextBox.Text;
            if (fileToConvert == null || fileToConvert.Equals(""))
            {
                MessageBox.Show("Select a file to convert");
            }
            else
            {
                if (!IsExcelInstalled())
                {
                    MessageBox.Show("Excel is not properly installed");
                }
                else
                {
                    backgroundWork = new BackgroundWork(fileToConvert);
                    backgroundWork.Progress += new BackgroundWork.ProgressDelegate(DisplayProgess);
                    backgroundWork.MakeUIEnabled += new BackgroundWork.EnableUI(EnableUIControls);
                    backgroundWork.AddItem += new BackgroundWork.AddSkippedItems(AddSkippedItem);

                    if (!backgroundWork.myConvertor.IsBusy)
                    {
                        ConvertToExcelButton.Enabled = false;
                        backgroundWork.myConvertor.RunWorkerAsync(xlApp);
                    }
                    else
                    {
                        //TODO : what to do when thread is busy
                    }
                }
            }
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            SkippedItemsList.Items.Clear();
            DialogResult result = TextFileDialog.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                TextFilePathtextBox.Text = TextFileDialog.FileName;
            }
            else
            {

            }
        }

        public bool IsExcelInstalled()
        {
            try
            {
                xlApp = new Excel.Application();
                if (xlApp == null)
                    return false;
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }

        }

        public SaveFileDialog GetSaveFileDialog()
        {
            return SaveExcelFileDialog;
        }

        public void DisplayProgess( int percent)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new BackgroundWork.ProgressDelegate(DisplayProgess), new Object[] {percent });
            }
            else
            {
                this.toolStripProgressBar1.Value = percent;
            }
        }

        public void EnableUIControls(bool shdEnable)
        {
            ConvertToExcelButton.Enabled = shdEnable;
            TextFilePathtextBox.Text = "";

        }

        public void AddSkippedItem(String str)
        {
            SkippedItemsList.Items.Add(str);
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {

        }

        private void ClearList_Click(object sender, EventArgs e)
        {
            SkippedItemsList.Items.Clear();
        }
    }
}

