using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace ExcelComparator.Test
{
    class Tets
    {
       public void testMain()
        {
            ExcelComparator.Parser.ExcelParser parser = new Parser.ExcelParser();
            parser.openFile(@"C:\Users\vivek_\Documents\Visual Studio 2010\Projects\ExcelComparator\DE_Economy_Quarterly_QRY.xlsx");
            int ret = parser.selectSheet(2);
            if ( ret < 0 )
            {
                Console.WriteLine("Error in opening sheet " + ret);
                return;
            }
            Console.WriteLine("Printing rows");

            List<string> row;
            while ( (row = parser.getNextRow()).Count > 0 )
            {
                  
            }
        }

       public void testComp()
       {
           ExcelComparator.Comparator.ExcelComparetor ecmp = new Comparator.ExcelComparetor();
           ecmp.openFiles(@"C:\Users\vivek_\Documents\Visual Studio 2010\Projects\ExcelComparator\f1.xlsx",
               @"C:\Users\vivek_\Documents\Visual Studio 2010\Projects\ExcelComparator\f2.xlsx");
           List<Comparator.DiffCell> diff =  ecmp.findDiff();
           Console.WriteLine("Diff count is " + diff.Count);
       }
    }
}
