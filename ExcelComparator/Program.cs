using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ExcelComparator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Test.Tets ts = new Test.Tets();
            ts.testComp();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
