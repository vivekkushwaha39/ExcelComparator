using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelComparator.Parser
{
    class ExcelParser
    {
        private Excel.Workbook wrkBook = null ;
        private Excel._Worksheet selectedSheet = null ;
        private int currCol = -1;
        private int currRow = -1;

        private Excel.Range usedRange;

        public void close()
        {

            wrkBook.Close();
            Marshal.ReleaseComObject(selectedSheet);
            Marshal.ReleaseComObject(wrkBook);
            GC.Collect();
        }

        public void openFile(string name)
        {
            Excel.ApplicationClass appClass = new Excel.ApplicationClass();
            wrkBook = appClass.Workbooks.Open(name);
        }

        public int getTotalColCount()
        {
            if (selectedSheet == null)
            {
                return -1;
            }

            return usedRange.Columns.Count;
        }

        public int getTotalRowCount()
        {
            if (selectedSheet == null)
            {
                return -1;
            }

            return usedRange.Rows.Count;
        }

        public int selectSheet(int index)
        {
            if (wrkBook == null)
                return -1;

            Console.WriteLine("Numer of sheets are " + wrkBook.Sheets.Count );
            if ( index >= wrkBook.Sheets.Count )
            {
                return -2;
            }

            selectedSheet = (Excel._Worksheet)wrkBook.Sheets[index];
            usedRange = selectedSheet.UsedRange;
            currRow = 1;
            return 0;
        }

        public void printData()
        {
            Console.WriteLine(usedRange.Value2.ToString());
        }

        public void resetRowPointer()
        {
            currRow = 1;
        }

        public List<string>  getNextRow()
        {
            List<string> rowData = new List<string>();
            int numCols = usedRange.Columns.Count;
            if ( currRow <= usedRange.Rows.Count )
            {
                if (usedRange.Cells[currRow, numCols] != null)
                {
                    for (currCol = 1; currCol <= numCols; currCol++)
                    {
                        Excel.Range currRange = (Excel.Range)usedRange.Cells[currRow, currCol];
                        Console.Write("," + currRange.Value2 );
                        rowData.Add(currRange.Value.ToString());
                        Marshal.ReleaseComObject(currRange);
                    }
                    Console.WriteLine("");
                    currRow++;
                }
            }

            return rowData;
        }
    }
}
