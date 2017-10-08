using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace StringCal
{
  public  class SlateCodeTest
    {
      public void getSlateCodeTest_File()
        {

            //Creation of Component Object Model(COM) Objects that is reference a file
            Excel.Application Excelwork = new Excel.Application();
            Excel.Workbook ExcelWorkbook = Excelwork.Workbooks.Open(@"C:\Test\slatecode");
            Excel._Worksheet ExcelWorksheet = ExcelWorkbook.Sheets[1];
            Excel.Range myRange = ExcelWorksheet.UsedRange;


          // declaration of rows and colums Count 
            int myRowCount = myRange.Rows.Count;
            int myColCount = myRange.Columns.Count;

            //iterate over the rows and columns and print to the console 
            //excel cannot be zero based, that is while i is initialized to 1 (int i = 1) also while
            for (int i = 1; i <= myRowCount; i++)
            {
                for (int j = 1; j <= myColCount; j++)
                {                    
                    if (j == 1)//new line
                        Console.Write("\r\n");

                    //write the value to the console
                    if (myRange.Cells[i, j] != null && myRange.Cells[i, j].Value2 != null)
                        Console.Write(myRange.Cells[i, j].Value2.ToString() + "\t");
                    
                }
            }

            //Garbage Collection
            GC.Collect();
            GC.WaitForPendingFinalizers();
          
            //release com objects to fully kill excel process from running after execution 
            Marshal.ReleaseComObject(myRange);
            Marshal.ReleaseComObject(ExcelWorksheet);

            //close the workbook 
            ExcelWorkbook.Close();
            Marshal.ReleaseComObject(ExcelWorkbook);

            //quit and release
            Excelwork.Quit();
            Marshal.ReleaseComObject(Excelwork);
        }
    }
    
}
