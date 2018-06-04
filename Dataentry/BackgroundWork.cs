using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Dataentry
{
    class BackgroundWork
    {
        public delegate void ProgressDelegate(int percent);
        public event ProgressDelegate Progress;
        public delegate void EnableUI(bool shdEnable);
        public delegate void AddSkippedItems(String str);
        public event AddSkippedItems AddItem;
        public event EnableUI MakeUIEnabled;
        public BackgroundWorker myConvertor;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Application xlApp;
        object misValue;
        String fileToConvert;
        List<String> itemsList;
        MainWindow mainWindow;
        public BackgroundWork(String fileToConvert)
        {
            myConvertor = new BackgroundWorker();
            myConvertor.DoWork += new DoWorkEventHandler(MyConvertor_DoWork);
            myConvertor.RunWorkerCompleted += new RunWorkerCompletedEventHandler(MyConvertor_RunWorkerCompleted);
            myConvertor.ProgressChanged += new ProgressChangedEventHandler(MyConvertor_ProgressChanged);
            myConvertor.WorkerReportsProgress = true;
            myConvertor.WorkerSupportsCancellation = false;
            itemsList = new List<String>();
            this.fileToConvert = fileToConvert;
            mainWindow = new MainWindow();
        }

        private void MyConvertor_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker sendingWorker = (BackgroundWorker)sender;//Capture the BackgroundWorker that fired the event
            xlApp = (Excel.Application)e.Argument;

            misValue = System.Reflection.Missing.Value;

            try{
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //Pass the file path and file name to the StreamReader constructor
                StreamReader sr = new StreamReader(fileToConvert);

                //Console.Write("Do work");
                String[] columns = {"SR.NO.", "NAME", "PAN NUMBER", "RANK", "GROSS SALARY", "DEDUCTION", "TOTAL INCOME", "FORCE No.","DCPS", "TAX PAYABLE", "TAX DEDUCTED" };
                
                for (int j = 0; j < columns.Length; j++)
                {
                    xlWorkSheet.Cells[1, j + 1] = columns[j];
                }
                DecodeTextFile(sr, sendingWorker);
                //close the file
                sr.Close();
            }
            catch(Exception exception)
            {
               
            }  
        }

        private void MyConvertor_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                try
                {
                    SaveFileDialog saveFileDialog = mainWindow.GetSaveFileDialog();
                    DialogResult result = saveFileDialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        try {
                            xlWorkSheet.Rows.AutoFit();
                            xlWorkSheet.Columns.AutoFit();
                            xlWorkSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                            xlWorkBook.SaveAs(saveFileDialog.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                            for(int i=0;i<itemsList.Count;i++)
                                AddItem?.Invoke((i+1)+". "+itemsList[i]);
                            MessageBox.Show("Excel file created");
                        }
                        catch(Exception exception)
                        {
                            String msg = exception.Message.ToString();
                            if (msg.StartsWith("Cannot access"))
                                MessageBox.Show("Excel file is already open. Please close it and try again");
                            else if (msg.Equals("Exception from HRESULT: 0x800A03EC"))
                                MessageBox.Show("Excel file not replaced");
                            else
                                MessageBox.Show(msg);
                        }
                        
                    }
                    else
                    {
                        MessageBox.Show("Excel file not created");
                    }

                    xlWorkBook.Close(false, misValue, misValue);
                    xlApp.Quit();
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                }
            }
            else
            {
                MessageBox.Show(e.ToString());
            }

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            MakeUIEnabled?.Invoke(true);
            Progress?.Invoke(0);

        }
        public void MyConvertor_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Progress?.Invoke(e.ProgressPercentage);
        }

        private void DecodeTextFile(StreamReader sr, BackgroundWorker sendingWorker)
        {
            String line = sr.ReadLine();
            line = line.Trim();
            int totalLines = File.ReadAllLines(fileToConvert).Length;
            int i = 1;
            int lineNum = 1;

            //Continue to read until you reach end of file
            while (line != null)
            {
                String result;
                if (line.Length > 0)
                {

                    if (line.StartsWith("Name"))
                    {
                        i++;
                        int pFrom = line.IndexOf("Name and Address of the Employer  Name:-") + "Name and Address of the Employer  Name:-".Length;
                        int pTo = line.IndexOf("Page");
                        int len = pTo - pFrom;
                        if (len > 0 && pFrom >= 0)
                        {
                            result = line.Substring(pFrom, len);
                            result = result.Trim();
                        }
                        else
                        {
                            pFrom = line.IndexOf("Name and Address of the Employer  Name:-") + "Name and Address of the Employer  Name:-".Length;
                            pTo = line.IndexOf("age");
                            len = pTo - pFrom;
                            
                            if (len > 0 && pFrom >= 0)
                            {
                                result = line.Substring(pFrom, len);
                                result = result.Trim();
                            }
                            else
                            {
                                result = "";
                                itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                            }
                        }
                        xlWorkSheet.Cells[i, 1] = i - 1;
                        xlWorkSheet.Cells[i, 2] = result;
                    }

                    else if (line.StartsWith("Force"))
                    {
                        int pFrom = line.IndexOf("Force No:") + "Force No:".Length;
                        int pTo = line.IndexOf("Pan");
                        int len = pTo - pFrom;
                        if (len > 0 && pFrom >= 0)
                        {
                            result = line.Substring(pFrom, len);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 8] = result;

                        pFrom = line.IndexOf("Pan No:-") + "Pan No:-".Length;
                        if (pFrom >= 0)
                        {
                            result = line.Substring(pFrom);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 3] = result;

                    }
                    else if (line.StartsWith("8."))
                    {
                        int pFrom = line.IndexOf("8. GROSS TOTAL INCOME (6+7)") + "8. GROSS TOTAL INCOME (6+7)".Length;
                        if (pFrom >= 0)
                        {
                            result = line.Substring(pFrom);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 5] = result;

                    }

                    else if (line.StartsWith("DCPS matching Govt"))
                    {
                        int pFrom = line.IndexOf("DCPS matching Govt.Contribution") + "DCPS matching Govt.Contribution  ".Length;
                        if (pFrom >= 0)
                        {
                            result = line.Substring(pFrom);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 9] = result;

                    }
                    


                    else if (line.StartsWith("10."))
                    {
                        int pFrom = line.IndexOf("10. Aggregate of deductible amount") + "10. Aggregate of deductible amount".Length;
                        if (pFrom >= 0)
                        {
                            result = line.Substring(pFrom);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 6] = result;

                    }

                    else if (line.StartsWith("11."))
                    {
                        int pFrom = line.IndexOf("11. TOTAL INCOME (8-10)") + "11. TOTAL INCOME (8-10)".Length;
                        int pTo = line.IndexOf("or");
                        int len = pTo - pFrom;
                        if (len > 0 && pFrom >= 0)
                        {
                            result = line.Substring(pFrom, len);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 7] = result;
                    }
                    else if (line.Contains("Rank"))
                    {
                        int pFrom = line.IndexOf("Rank:-") + "Rank:-".Length;
                        if (pFrom >= 0)
                        {
                            result = line.Substring(pFrom);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 4] = result;
                    }

                    else if (line.StartsWith("20"))
                    {
                        int pFrom = line.IndexOf("20.TAX PAYABLE") + "20.TAX PAYABLE  ".Length;
                        if (pFrom >= 0)
                        {
                            result = line.Substring(pFrom);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 10] = result;

                    }

                    else if (line.StartsWith("21."))
                    {
                        int pFrom = line.IndexOf("21. LESS :(a)Tax deducted at source u/s192(I)") + "21. LESS :(a)Tax deducted at source u/s192(I)".Length;
                        if (pFrom >= 0)
                        {
                            result = line.Substring(pFrom);
                            result = result.Trim();
                        }
                        else
                        {
                            result = "";
                            itemsList.Add("Sr.No " + (i - 1) + " Name Column");
                        }
                        xlWorkSheet.Cells[i, 11] = result;

                    }

                }
                
               //Read the next line
               line = sr.ReadLine();
                if (line != null)
                    line = line.Trim();
                int per = (lineNum * 100) / totalLines;
               
                sendingWorker.ReportProgress(per);//Report our progress to the main thread
                lineNum++;
            }
        }
    }
}
