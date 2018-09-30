using System;
using System.Text;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ReadPDFForm
{

    public partial class Form1 : Form
    {

        public int driverRow = 0;
        public int driverColumn = 0;

        public int rateRow = 0;
        public int rateColumn = 1;

        Excel xl = new Excel("C:\\test.xlsx", 1); //<==== Enter Excel location here

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.Multiselect = true;
            dlg.Filter = "PDF Files(*.PDF)|*.PDF|All Files(*.*)|*.*";

            if (dlg.ShowDialog() == DialogResult.OK)
            {

                foreach (String file in dlg.FileNames)
                {

                    string strText = string.Empty;

                    try
                    {
                        PdfReader reader = new PdfReader(file);
                        for (int page = 1; page <= reader.NumberOfPages; page++)
                        {
                            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
                            string s = PdfTextExtractor.GetTextFromPage(reader, page, its);

                            s = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(s)));
                            strText = strText + s;

                            richTextBox1.Text = strText;

                        }

                        //<======== Adjust string location below ========>

                        int driverIndex = strText.IndexOf("Driver Pay Report") + 18;

                        string driver = strText.Substring(driverIndex); //Sets driver equal to substring after index.
                        driver = driver.Remove(driver.IndexOf("Address") - 1); //removes the rest of the string.

                        driver = driver.Remove(driver.IndexOf("Driver: "), 8); //Removes "Driver: " from string


                        int rateIndex = strText.IndexOf("Grand Total:") + 14;

                        string rate = strText.Substring(rateIndex); //Sets rate equal to substring after index.
                        rate = rate.Remove(rate.IndexOf(" USD"));//removes the rest of the string.

                        reader.Close();

                        xl.WriteToCell(driverRow, 0, driver); //Column is 0

                        xl.WriteToCell(rateRow, 1, rate); //Column is 1

                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + file);
                    }


                    driverRow++;
                    rateRow++;
                }

                xl.Save();
                xl.Close();

            }

        }

    }


    class Excel
    {

        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel()
        {

        }

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];

        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        }

        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }

        public void Save()
        {
            wb.Save();
        }

        public void Close()
        {
            wb.Close();
        }

    }

}