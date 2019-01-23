using System.Collections.Generic;
using System.Linq;
using System.Text;
using System;

namespace ExcelComparator.Comparator
{
    class ExcelComparetor
    {
        private Parser.ExcelParser file1;
        private Parser.ExcelParser file2;

        public void openFiles(string f1, string f2)
        {
            file1 = new Parser.ExcelParser();
            file2 = new Parser.ExcelParser();

            file1.openFile(f1);
            file1.selectSheet(2);
            
            file2.openFile(f2);
            file2.selectSheet(2);


        }

        public List<DiffCell> findDiff()
        {
            List<DiffCell> diff = new List<DiffCell>();

            

            if (file1.getTotalColCount() != file2.getTotalColCount() ||
                file1.getTotalRowCount() != file2.getTotalRowCount())
            {
                Console.WriteLine("Unable to compare row or columns not match");
            }
            int row = file1.getTotalRowCount();
            int col = file1.getTotalColCount();

            for (int i = 0; i < row; i++)
            {
                
                List<string> rowf1 = file1.getNextRow();
                List<string> rowf2 = file2.getNextRow();
                for (int j = 0; j < col; j++)
                {
                    if (rowf1.ElementAt(j) != rowf2.ElementAt(j))
                    {
                        DiffCell diffcell = new DiffCell();
                        diffcell.dataf1 = rowf1.ElementAt(j);
                        diffcell.dataf2 = rowf2.ElementAt(j);
                        diffcell.row = i;
                        diffcell.col = j;
                        diff.Add(diffcell);
                    }
                }

            }
            file1.close();
            file2.close();
            return diff;
        }

    }

    class DiffCell
    {
        public int row;
        public int col;
        public string dataf1;
        public string dataf2;
    }
}
