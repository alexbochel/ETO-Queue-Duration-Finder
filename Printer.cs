using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ETOCBurgDuration
{
    public class Printer
    {
        _Application excel;
        Workbooks wbs;
        Workbook wb;
        Worksheet ws;
        List<Data> masterList;
        const int sheet = 1;

        public Printer(List<Data> masterList)
        {
            excel = new _Excel.Application();
            excel.Visible = true;
            wbs = excel.Workbooks;
            wb = wbs.Add();
            this.ws = wb.Worksheets[sheet];
            this.masterList = masterList;

            print();
        }

        public void print()
        {
            printHeaders();
            printData();
        }

        private void printHeaders()
        {
            printCell(1, 1, "Sales Number");
            printCell(1, 2, "Material");
            printCell(1, 3, "Description");
            printCell(1, 4, "Created On");
            printCell(1, 5, "Days Available");
            printCell(1, 6, "Days Not Available");

            var range = ws.get_Range("A1", "F1");
            range.Font.Bold = true;
        }

        private void printData()
        {
            int horiz = 1;
            int buffer = 2;
            
            for (int i = 0; i < masterList.Count; i++ )
            {
                int row = i + buffer;
                horiz = 1;

                printCell(row, horiz, masterList[i].salesNum);
                horiz++;
                printCell(row, horiz, masterList[i].mat);
                horiz++;
                printCell(row, horiz, masterList[i].desc);
                horiz++;
                printCell(row, horiz, masterList[i].dateCreated);
                horiz++;
                printCell(row, horiz, masterList[i].daysAvailable.ToString());
                horiz++;
                printCell(row, horiz, masterList[i].daysNotAvailable.ToString());
                horiz++;
            }

            ws.Columns.AutoFit();
        }

        /// <summary>
        /// This method prints data in a cell. 
        /// </summary>
        /// <param name="i"> The "y" coordinate on a plane. </param>
        /// <param name="j"> The "x" coordinate on a plane. </param>
        /// <param name="value"> The data to be printed in the cell. </param>
        private void printCell(int i, int j, string value)
        {
            ws.Cells[i, j].Value2 = value;
        }
    }
}
