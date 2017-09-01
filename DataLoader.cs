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
    /// <summary>
    /// This class loads the worksheet into the current view of the program so that it can be worked on. 
    /// It does this for each file path given to it by the Program class. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 8/30/2017
    /// 
    /// </summary>
    public class DataLoader
    {
        /// <summary>
        /// Final list of ETO's for the desired period of time. 
        /// </summary>
        public List<Data> masterList;

        private string[] _excelPaths;
        private _Application _excel;
        private Workbooks _wbs;
        private Worksheet ws;
        private Finder finder;
        private Printer printer;

        /// <summary>
        /// G/S: The array of paths to all of the files in the folder selected by the user. 
        /// </summary>
        public string[] excelPaths
        {
            get { return _excelPaths; }
            set { _excelPaths = value; }
        }

        /// <summary>
        /// G/S: Excel instance in use. 
        /// </summary>
        public _Application excel
        {
            get { return _excel; }
            set { _excel = value; }
        }

        /// <summary>
        /// G/S: Workbooks in use. 
        /// </summary>
        public Workbooks wbs
        {
            get { return _wbs; }
            set { _wbs = value; }
        }

        /// <summary>
        /// This constructor loads the paths from the Program class and creates new instances of
        /// excel, a list, and a finder. 
        /// </summary>
        /// <param name="pathArray"></param>
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

        /// <summary>
        /// This method ensures that the worksheet being looked at is the sheet with the necessary data. 
        /// </summary>
        /// <param name="wb"> Workbook in which a sheet is being looked for. </param>
        private void findWorksheet(Workbook wb)
        {
            for (int i = 1; i <= wb.Worksheets.Count; i++)
            {
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

        /// <summary>
        /// Disconnects all excel instances from the computer memory. 
        /// </summary>
        private void garbageCleanup()
        {
            wbs.Close();
            excel.Quit();

            Marshal.ReleaseComObject(wbs);
            Marshal.ReleaseComObject(excel);
        }

    }
}
