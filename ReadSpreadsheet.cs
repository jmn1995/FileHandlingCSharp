using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Windows.Forms;
using System.Data;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using FileHandlingProgram;

namespace FileHandlingProgram
{
 
        class ReadSpreadsheet
        {
        List<ActivityModel> ExcelSchedule = new List<ActivityModel>();
            public List<ActivityModel> ReadExcelFile(string filepath)
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filepath);
                Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                object[,] values = (Object[,])xlRange.Value2;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
               
                //Start at row two to skip header
                for (int rowNum = 1; rowNum <= rowCount; rowNum++)
                {

                    ActivityModel FillBox = new ActivityModel();

                    for (int colNum = 1; colNum <= colCount; colNum++)
                    {
                    
                    Console.Write(Convert.ToString(values[rowNum, colNum]) + " | " );


                        if (colNum == 1)
                        {
                         FillBox = new ActivityModel();
                        }
                        if (values[rowNum, colNum] != null)
                        {
                            switch (colNum)
                            {
                                case 1:
                                    FillBox.ActivityName = Convert.ToString(values[rowNum, colNum]);
                                    break;
                                case 2:
                                    FillBox.Monday = Convert.ToString(values[rowNum, colNum]);
                                    break;
                                case 3:
                                    FillBox.Tuesday = Convert.ToString(values[rowNum, colNum]);
                                    break;
                                case 4:
                                    FillBox.Wednesday = Convert.ToString(values[rowNum, colNum]);
                                    break;

                                default:
                                    break;
                            }
                    
                        }
                    }//end inner 'for' loop
                    Console.WriteLine("");
                    ExcelSchedule.Add(FillBox);
                }
                xlWorkbook.Close();
                xlApp.Quit();
            return ExcelSchedule;
            }
        }
    }