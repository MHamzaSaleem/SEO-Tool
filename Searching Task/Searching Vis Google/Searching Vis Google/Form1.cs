using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;
using OpenQA.Selenium.Chrome;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using OpenQA.Selenium;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop;

namespace Searching_Vis_Google
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            Thread thread = new Thread(new ThreadStart(StartForm));
            thread.Start();
            Thread.Sleep(5000);
            InitializeComponent();
            thread.Abort();
        }

        public void StartForm()
        {
            try
            {
                System.Windows.Forms.Application.Run(new AppStart());
            }
            catch(Exception ex)
            {
                //MessageBox.Show("Welcome To Searching Engine");
            }
        }
        string[] res = new string[2];
        public static int len = 0;
        public static int currentLine=0;
        string[] searchlines = new string[0];
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Data.xlsx";
        object selectedValue = "";
        public static int linenumber = 0;
        List<Excel> gridViewList = new List<Excel>();
        public static string[,] finalResult = new string[0, 3];
        public static int arrayLen = 0;
        
        private void timer2_Tick(object sender, EventArgs e)
        {
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            if (finalResult.Length != gridViewList.Count)
            {
                var source = new BindingSource();
                source.DataSource = gridViewList;
                dataGridView1.DataSource = source;
            }
            else
                timer2.Stop();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Visible = label3.Visible = false;  
        }


        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            gridViewList.Clear();
            RefreshTimer();
            if (File.Exists(path))
            {
                label1.Text = "";
                len = richTextBox1.Lines.Length;
                if (!richTextBox1.Text.Equals("") && comboBox1.SelectedIndex > -1)
                {
                    Array.Resize(ref  searchlines, searchlines.Length + len);
                    for (int i = 0; i < len; i++)
                    {
                        try
                        {
                            searchlines[i] = richTextBox1.Lines[i];
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    progressBar1.Visible = label3.Visible = true;
                    progressBar1.Style = ProgressBarStyle.Marquee;
                    button1.Enabled = richTextBox1.Enabled = comboBox1.Enabled = false;
                    toCheck();
                }
                else if (comboBox1.SelectedIndex == -1)
                {
                    MessageBox.Show("Please Select Any Country First!");
                    label1.Text = "No Result found!";
                }
                else
                {
                    MessageBox.Show("Search Box is Empty!");
                     label1.Text = "No Result found!";
                }
            }
            else
            {
                MessageBox.Show("Bofore Running this Application You Have To Create Excel File On Desktop With The Name Of \'Data\' & Its Extenion Should Be .xlsx Thank You For Your Time!");
            }
        }

        private void RefreshTimer()
        {
            timer2.Interval = (3 * 1000); // 3 secs
            timer2.Tick += new System.EventHandler(timer2_Tick);
            timer2.Start();
        }

        private async void toCheck()
        {
            await Task.Run(() => Search());
            if (currentLine == linenumber)
                FinalStep();
            else
                toCheck();
        }

        private void FinalStep()
        {
            searchlines = new string[0];
            Array.Resize(ref  searchlines, searchlines.Length + len);
            currentLine = linenumber = 0;
            progressBar1.Visible = label3.Visible = false;       
            button1.Enabled = richTextBox1.Enabled = comboBox1.Enabled = true;
            richTextBox1.Text = "";
            comboBox1.SelectedItem = null;
            label1.Text = "Data Have Been Saved Successfully!";
            timer1.Interval = 5000; 
            timer1.Tick += (s, e) =>
            {
                label1.Text="";
                timer1.Stop();
            };
            timer1.Start();
        }

        private void Search()
        {
            try
            {
                if (currentLine == 0)
                    finalResult = new string[len, 3];
                string searchFilter = "allintitle:";
                string url = "";
                string ipport = "";
                string searchLoc = selectedValue.ToString();
                switch (searchLoc)
                {
                    case "UK":
                        url = "https://www.google.co.uk/search?gl=gb&q=" + searchFilter;
                        ipport = "185.10.166.130:8080";
                        break;

                    case "USA":
                        url = "https://www.google.com/search?gl=us&hl=en&pws=0&gws_rd=cr&q=" + searchFilter;
                        ipport = "138.68.240.218:8080";
                        break;

                    case "Pakistan":
                        url = "https://google.com.pk/search?q=" + searchFilter;
                        ipport = "125.209.116.14:8080";
                        break;

                    case "Canada":
                        url = "https://www.google.ca/search?hl=fr&gws_rd=ssl&q=" + searchFilter;
                        ipport = "74.15.191.160:41564";
                        break;

                    default:
                        url = "http://google.com/search?q=" + searchFilter;
                        ipport = "localhost:8888";
                        break;
                }

                var ser = ChromeDriverService.CreateDefaultService();
                ser.HideCommandPromptWindow = true;
                var chromeOptions = new ChromeOptions();
                var proxy = new Proxy();
                proxy.HttpProxy = ipport;
                chromeOptions.Proxy = proxy;
                chromeOptions.AddArguments("headless");
                using (var driver = new ChromeDriver(ser, chromeOptions))
                {
                    try
                    {
                        for (int i = currentLine; i < len; i++)
                        {
                            currentLine++;
                            driver.Navigate().GoToUrl(url + searchlines[i]);

                            var getElement = driver.FindElementById("mBMHK");
                            string getText = getElement.GetAttribute("innerHTML");
                            //thisarray
                            res = getText.Split(' ');
                            if (searchLoc == "Canada")
                            {
                                res[0] = "About ";
                                res[1] = Regex.Replace(res[1], "[^0-9]+", string.Empty);
                            }
                            
                            finalResult[i, 0] = searchlines[i];
                            finalResult[i, 1] = res[0];
                            finalResult[i, 2] = res[1];
                            linenumber = finalResult.GetLength(0);
                            gridViewList.Add(new Excel { Phrase = finalResult[i, 0], Results = finalResult[i, 1] + " " + finalResult[i, 2] });
                        }
                        driver.Close();
                    }
                    catch (Exception ex)
                    {
                        int line = currentLine - 1;
                        if (searchlines[line].Equals(""))
                        {
                            finalResult[line, 0] = searchlines[line] + "Empty line";
                            finalResult[line, 1] = "Empty Line";
                            finalResult[line, 2] = "";
                        }
                        else
                        {
                            finalResult[line, 0] = searchlines[line];
                            finalResult[line, 1] = "Result Not Found";
                            finalResult[line, 2] = "";
                        }
                        gridViewList.Add(new Excel { Phrase = finalResult[line,0], Results = finalResult[line, 1] + " " + finalResult[line, 2] });
                    }
                    if (currentLine == linenumber)
                        saveToExcel(finalResult);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void saveToExcel(string[,] finalResult)
        {
            try
            {
                if (!File.Exists(path))
                {
                    Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = false;
                    Microsoft.Office.Interop.Excel._Workbook oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                    Microsoft.Office.Interop.Excel._Worksheet oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                    oSheet.Cells[1, 1] = "Phrase";
                    oSheet.Cells[1, 2] = "Results";
                    oSheet.Rows[1].Cells[1].Interior.Color = System.Drawing.Color.OrangeRed;
                    oSheet.Rows[1].Cells[2].Interior.Color = System.Drawing.Color.OrangeRed;
                    oSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                    for (int i = 2; i < finalResult.GetLength(0) + 2; i++)
                    {
                        oSheet.Cells[i, 1] = finalResult[i - 2, 0];
                        oSheet.Cells[i, 2] = finalResult[i - 2, 1] + " " + finalResult[i - 2, 2];
                    }

                    oSheet.Columns.AutoFit();
                    oXL.DisplayAlerts = false;
                    oXL.Visible = false;
                    oXL.UserControl = false;
                    oWB.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    oWB.Close();
                    oXL.Quit();
                    Marshal.ReleaseComObject(oSheet);

                    Marshal.ReleaseComObject(oWB);

                    Marshal.ReleaseComObject(oXL);
                }
                else
                {
                    Excel a = new Excel();
                    _Application xlApp = new _Excel.Application();
                    Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                    Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Range xlRange = xlWorksheet.UsedRange;
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count - 1;
                    List<Excel> list = new List<Excel>();
                    if (rowCount > 1)
                    {
                        for (int i = 1; i <= rowCount; i++)
                        {
                            string phrase = "";
                            string result = "";
                            for (int j = 1; j <= colCount + 1; j++)
                            {
                                try
                                {
                                    if (j == 1)
                                    {
                                        phrase = xlRange.Cells[i, j].Value2.ToString();
                                    }
                                    else
                                    {
                                        result = xlRange.Cells[i, j].Value2.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }
                            }
                            list.Add(new Excel { Phrase = phrase, Results = result });
                        }
                    }
                    else
                    {
                        list.Add(new Excel { Phrase = "Phrase", Results = "Result" });
                    }
                    for (int i = rowCount; i < finalResult.GetLength(0) + rowCount; i++)
                    {
                        list.Add(new Excel { Phrase = finalResult[i - rowCount, 0], Results = finalResult[i - rowCount, 1] + " " + finalResult[i - rowCount, 2] });
                    }
                    xlWorksheet.Rows[1].Cells[1].Interior.Color = System.Drawing.Color.OrangeRed;
                    xlWorksheet.Rows[1].Cells[2].Interior.Color = System.Drawing.Color.OrangeRed;
                    xlWorksheet.Cells[1, 1].EntireRow.Font.Bold = true;
                    int lineNumber = 1;
                    foreach (var items in list)
                    {
                        xlWorksheet.Cells[lineNumber, 1] = items.Phrase;
                        xlWorksheet.Cells[lineNumber, 2] = items.Results;
                        lineNumber++;
                    }

                    xlWorksheet.Columns.AutoFit();
                    xlApp.Visible = false;
                    xlApp.UserControl = false;
                    xlApp.DisplayAlerts = false;
                    xlWorkbook.Save();
                    xlWorkbook.Close(true);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkbook);

                    Marshal.ReleaseComObject(xlWorkbook);

                    Marshal.ReleaseComObject(xlApp);
                    

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                    process.Kill();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedValue = comboBox1.SelectedItem;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
       {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void backgroundWorker1_DoWork_1(object sender, DoWorkEventArgs e)
        {

        }


       

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}