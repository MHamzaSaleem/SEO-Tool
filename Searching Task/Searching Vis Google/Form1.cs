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
using OpenQA.Selenium.Chrome;
using System.IO;


namespace Searching_Vis_Google
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        object selectedValue = "";
        string finalResult = "";
        private void Form1_Load(object sender, EventArgs e)
        {
             comboBox1.SelectedItem = null;
             comboBox1.SelectedText = "--Select Country--";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!richTextBox1.Text.Equals(""))
            {
                Search(richTextBox1.Text);
            }
            else
            {
                MessageBox.Show("Search Box is Empty!");
                label1.Text = "No Result found!";
            }
        }

        private void Search(string search)
        {
            label1.Text = "";
            string searchFilter = "allintitle:" + search;
            string url = "";
            string searchLoc = selectedValue.ToString();
            switch (searchLoc)
            {
                case "UK":
                    url = "http://google.co.uk/search?q=" + searchFilter;
                    break;

                case "USA":
                    url = "https://www.google.com/search?gl=us&hl=en&pws=0&gws_rd=cr&q=" + searchFilter;
                    break;

                case "Pakistan":
                    url = "http://google.com.pk/search?q=" + searchFilter;
                    break;

                //case "Canada":
                //    url = "https://www.google.ca/search?hl=fr&gws_rd=ssl&q=" + searchFilter;
                //    break;

                case "Global":
                    url = "http://google.com/search?q=" + searchFilter;
                    break;

                default:
                    url = "http://google.com/search?q=" + searchFilter;
                    break;
            }          
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");

            using (var driver = new ChromeDriver(chromeOptions))
            {
                try
                {
                    driver.Navigate().GoToUrl(url);

                    var getElement = driver.FindElementById("resultStats");
                    string getText = getElement.GetAttribute("innerHTML");
                    string[] res = getText.Split(' ');
                    finalResult =  res[0] + " " +res[1];
                    label1.Text = finalResult;
                    writeInFile(finalResult);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void writeInFile(string data)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+"\\SaveData.txt";
            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("Phrase" + "                               " + "Results");
                    
                    for (int n = 0; n < richTextBox1.Lines.Length; ++n)
                    {
                        if (n == 0)
                            sw.WriteLine(richTextBox1.Lines[n]+ "    " + data);
                        else
                            sw.WriteLine(richTextBox1.Lines[n]);
                    } 
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(path))
                {
                    for (int n = 0; n < richTextBox1.Lines.Length; ++n)
                    {
                        if (n == 0)
                            sw.WriteLine(richTextBox1.Lines[n] + "    " + data);
                        else
                            sw.WriteLine(richTextBox1.Lines[n]);
                    } 
                }	
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
    }
}
