using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ETOCBurgDuration
{
    public class DataLoader
    {
        public List<Data> masterList;

        private string[] _excelPaths;
        private _Application _excel;
        private Workbooks _wbs;
        private Worksheet ws;
        private Finder finder;
        private Printer printer;

        public string[] excelPaths
        {
            get { return _excelPaths; }
            set { _excelPaths = value; }
        }

        public _Application excel
        {
            get { return _excel; }
            set { _excel = value; }
        }

        public Workbooks wbs
        {
            get { return _wbs; }
            set { _wbs = value; }
        }

        public DataLoader(string[] pathArray)
        {
            excelPaths = pathArray;
            masterList = new List<Data>();
            finder = new Finder();
            excel = new _Excel.Application();
            wbs = excel.Workbooks;

            excel.Visible = false;
        }

        /// <summary>
        /// This method iterates through the paths to all of the excel files and
        /// creates a workbook for each one. 
        /// </summary>
        public void execute()
        {
            foreach(string book in excelPaths)
            {
                Workbook wb = wbs.Open(book);
                wb.UpdateLinks = XlUpdateLinks.xlUpdateLinksNever;
                
                findWorksheet(wb);

                finder.find(ws, ref masterList);

                wb.Close();
                Marshal.ReleaseComObject(wb);
            }

            wbs.Close();
            excel.Quit();
            printer = new Printer(masterList);
        }

        private void findWorksheet(Workbook wb)
        {
            for (int i = 1; i <= wb.Worksheets.Count; i++)
            {
                //try
                //{
                //    ws = wb.Worksheets[i];
                //}
                //catch
                //{
                //    Console.WriteLine("Checking sheets");
                //}

                ws = wb.Worksheets[i];

                bool a = (ws.Cells[17, 1].Text == "");
                bool b = (ws.Cells[1, 3].Text == "");
                bool c = (ws.Cells[17, 8].Text == "");
                bool d = (ws.Cells[1, 1].Text != "");

                if (a && b && c && d)
                {
                    break;
                }
            }
        }

        private void garbageCleanup()
        {
            wbs.Close();
            excel.Quit();

            Marshal.ReleaseComObject(wbs);
            Marshal.ReleaseComObject(excel);
        }

    }
}
