using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace FileHandlingProgram
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }
        string path = Environment.CurrentDirectory + "/" + "ABC.txt";
        private void button1_Click(object sender, EventArgs e)
        {
          
            if(!File.Exists(path))
   
                    {
                        StreamWriter writeText = File.CreateText(path);
                        MessageBox.Show("File created successfully");
                        writeText.Close();
                    }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            {
                using (StreamWriter sw = new StreamWriter(path))
            {
                sw.WriteLine("\n" + "This file can now be edited." + "\n");
                    MessageBox.Show("Text File has been written");
                    
      
            }
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
                MessageBox.Show("File has been deleted"); ;

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (StreamReader sr = new StreamReader(path))
            {
                string text = sr.ReadLine();
                richTextBox1.Text = text;
                
            }
        }

        private void button5_Click(object sender, EventArgs e)
          
           

        {
            var week = new List<Schedule>()
            {
                new Schedule() {Name ="Working", MondayHours =8, TuesdayHours =8, WednesdayHours = 8},
                new Schedule() {Name ="Excercising", MondayHours =1, TuesdayHours =1, WednesdayHours = 3},
                new Schedule() {Name ="Sleeping", MondayHours =6, TuesdayHours =7, WednesdayHours = 8}
                

            };

            CreateSpreadsheet(week);
            MessageBox.Show("Spreadsheet has been created");
        }

        public void CreateSpreadsheet(List<Schedule> week)
        {
            string SpreadsheetPath = "WeeklySchedule.xlsx";
            File.Delete(SpreadsheetPath);
            FileInfo SpreadsheetInfo = new FileInfo(SpreadsheetPath);

            ExcelPackage pck = new ExcelPackage(SpreadsheetInfo);
            var ScheduleWorksheet = pck.Workbook.Worksheets.Add("Schedule");
            ScheduleWorksheet.Cells["A1"].Value = "Name";
            ScheduleWorksheet.Cells["B1"].Value = "Monday";
            ScheduleWorksheet.Cells["C1"].Value = "Tuesday";
            ScheduleWorksheet.Cells["D1"].Value = "Wednesday";
            ScheduleWorksheet.Cells["A1:D1:"].Style.Font.Bold = true;

            int currentRow = 2;
            foreach (var sheet in week)
            {
                ScheduleWorksheet.Cells["A" + currentRow.ToString()].Value = sheet.Name;
                ScheduleWorksheet.Cells["B" + currentRow.ToString()].Value = sheet.MondayHours;
                ScheduleWorksheet.Cells["C" + currentRow.ToString()].Value = sheet.TuesdayHours;
                ScheduleWorksheet.Cells["D" + currentRow.ToString()].Value = sheet.WednesdayHours;

                currentRow++;

            }

            ScheduleWorksheet.View.FreezePanes(2, 1);

            ScheduleWorksheet.Cells["B" + (currentRow).ToString()].Formula = "SUM(B2:B" + (currentRow - 1).ToString() + ")";
            ScheduleWorksheet.Cells["C" + (currentRow).ToString()].Formula = "SUM(C2:C" + (currentRow - 1).ToString() + ")";
            ScheduleWorksheet.Cells["D" + (currentRow).ToString()].Formula = "SUM(D2:D" + (currentRow - 1).ToString() + ")";
            ScheduleWorksheet.Cells["B" + (currentRow).ToString()].Style.Font.Bold = true;
            ScheduleWorksheet.Cells["C" + (currentRow).ToString()].Style.Font.Bold = true;
            ScheduleWorksheet.Cells["D" + (currentRow).ToString()].Style.Font.Bold = true;
            ScheduleWorksheet.Cells["B" + (currentRow).ToString()].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ScheduleWorksheet.Cells["C" + (currentRow).ToString()].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ScheduleWorksheet.Cells["D" + (currentRow).ToString()].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            pck.Save();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ReadSpreadsheet tableCall = new ReadSpreadsheet();

            List<ActivityModel> ListPath = tableCall.ReadExcelFile(@"C:\Users\jnicholson\Desktop\FileHandlingProgram\FileHandlingProgram\bin\Debug\WeeklySchedule");
            foreach (ActivityModel a in ListPath)
            {
                string BlankName = a.ActivityName;
                if (BlankName == "" || BlankName == null )
                {
                    BlankName = "Total";
                }



                richTextBox1.Text += "\n" + "\n" + BlankName + " | " + a.Monday + " | " + a.Tuesday + " | " + a.Wednesday + System.Environment.NewLine;
                
            } 
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string SpreadsheetPath = "WeeklySchedule.xlsx";
            File.Delete(SpreadsheetPath);
            MessageBox.Show("Spreadsheet has been deleted");
            
        }

        private void button8_Click(object sender, EventArgs e)
        {
           
            string fname = "";
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            string ExcelFilePath = @"C: \Users\jnicholson\Desktop\FileHandlingProgram\FileHandlingProgram\bin\Debug\WeeklySchedule.xlsx";
            fdlg.InitialDirectory = new FileInfo(ExcelFilePath).DirectoryName;
            
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.FileName;
            }


            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // dt.Column = colCount;  
            dataGridView1.ColumnCount = colCount;
            dataGridView1.RowCount = rowCount;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {


                    //write the value to the Grid  


                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                    }
                    // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  

                    //add useful things here!     
                }
            }

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:  
            //  never use two dots, all COM objects must be referenced and released individually  
            //  ex: [somthing].[something].[something] is bad  

            //release com objects to fully kill excel process from running in the background  
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release  
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            using(SaveFileDialog SavePdf = new SaveFileDialog() {Filter="PDFFile|*.pdf", ValidateNames=true})
            {
                if (SavePdf.ShowDialog() == DialogResult.OK)
                {
                    iTextSharp.text.Document Doc = new iTextSharp.text.Document(PageSize.A4.Rotate());
                    try
                    {
                        PdfWriter.GetInstance(Doc, new FileStream(SavePdf.FileName, FileMode.Create));
                        Doc.Open();
                        Doc.Add(new iTextSharp.text.Paragraph(richTextBox1.Text));
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    finally
                    {
                        MessageBox.Show("File has been saved to pdf format");
                        Doc.Close();
                    }
                        
                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
