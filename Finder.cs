using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ETOCBurgDuration
{
    /// <summary>
    /// This class is given a worksheet and then looks through the sheet to find unique ETO's and to count the days 
    /// they are available/unavailable. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 8/30/2017
    /// 
    /// </summary>
    public class Finder
    {

        private Worksheet ws;

        /// <summary>
        /// This method searches the worksheet and compares the sales to those already in the master list. 
        /// </summary>
        /// <param name="ws"> Worksheet being searched. </param>
        /// <param name="masterList"> The master list of sales. </param>
        public void find(Worksheet ws, ref List<Data> masterList)
        {
            this.ws = ws;
            int row = 19;
            int column = 1; 

            // Iterate to the bottom of the Excel sheet. 
            while (readCell(column, row) != "")
            {
                int MRPCCol = 6;
                int salesNumCol = 1;
                int matCol = 2;
                int descCol = 3;
                int dateCol = 8;

                if (row > 19)
                {
                    if (readCell(matCol, row) == readCell(matCol, (row - 1)))
                    {
                        if (readCell(salesNumCol, row) != readCell(salesNumCol, (row - 1)))
                        {
                            Data newData = new Data();

                            newData.MRPC = readCell(MRPCCol, row);
                            newData.salesNum = readCell(salesNumCol, row);
                            newData.mat = readCell(matCol, row);
                            newData.desc = readCell(descCol, row);
                            newData.salesNum = readCell(salesNumCol, row);
                            newData.dateCreated = readCell(dateCol, row);

                            if (newData.dateCreated == "01/01/0001")
                            {
                                int Stop = 1;
                            }

                            checkMasterList(newData, ref masterList);
                        }
                    }
                    else
                    {
                        Data newData = new Data();

                        newData.MRPC = readCell(MRPCCol, row);
                        newData.salesNum = readCell(salesNumCol, row);
                        newData.mat = readCell(matCol, row);
                        newData.desc = readCell(descCol, row);
                        newData.salesNum = readCell(salesNumCol, row);
                        newData.dateCreated = readCell(dateCol, row);

                        if (newData.dateCreated == "01/01/0001")
                        {
                            int Stop = 1;
                        }

                        checkMasterList(newData, ref masterList);
                    }
                }
                else
                {
                    Data newData = new Data();

                    newData.MRPC = readCell(MRPCCol, row);
                    newData.salesNum = readCell(salesNumCol, row);
                    newData.mat = readCell(matCol, row);
                    newData.desc = readCell(descCol, row);
                    newData.salesNum = readCell(salesNumCol, row);
                    newData.dateCreated = readCell(dateCol, row);

                    checkMasterList(newData, ref masterList);
                }

                row++;
            }
        }

        /// <summary>
        /// Checks the master list for repeats and also records at what times the MRPC data indicates availability of the 
        /// ETO. 
        /// </summary>
        /// <param name="newData"> The new data just created from the worksheet. </param>
        /// <param name="masterList"> The master list of sales. </param>
        private void checkMasterList(Data newData, ref List<Data> masterList)
        {
            bool bFound = false;
            
            for (int i = 0; i < masterList.Count; i++)
            {
                if (newData.equals(masterList[i]))
                {
                    bFound = true;
                    
                    if (newData.MRPC == "BP3" && newData.mat.Contains("BEAC"))
                    {
                        masterList[i].daysAvailable++;
                        break;
                    }
                    else if (newData.MRPC == "ETO" && !newData.mat.Contains("BEAC"))
                    {
                        masterList[i].daysAvailable++;
                        break;
                    }
                    else
                    {
                        masterList[i].daysNotAvailable++;
                        break;
                    }
                }
            }

            if (!bFound)
            {
                if (newData.MRPC == "BP3" && newData.mat.Contains("BEAC"))
                {
                    newData.daysAvailable++;
                }
                else if (newData.MRPC == "ETO" && !newData.mat.Contains("BEAC"))
                {
                    newData.daysAvailable++;
                }
                else
                {
                    newData.daysNotAvailable++;
                }

                masterList.Add(newData);
            }
        }

        /// <summary>
        /// This method reads in a cell from excel. 
        /// </summary>
        /// <param name="i"> The "y" coordinate. </param>
        /// <param name="j"> The "x" coordinate. </param>
        /// <returns> The value in the cell as a string. </returns>
        private string readCell(int i, int j)
        {
            string cell = ws.get_Range(CellName(i, j), Type.Missing).Text.ToString();

            return cell;
        }

        // This method takes a number parameter in order to convert it to the correct Excel grid format. 
        private string ColumnName(int nColumn)
        {
            int tempCol, tempInt;
            string cellName, str;

            tempCol = nColumn;
            cellName = "";
            while (tempCol > 0)
            {
                tempInt = ((tempCol - 1) % 26) + 1;
                tempCol = (tempCol - tempInt) / 26;
                cellName += Convert.ToChar((tempInt + 64));
            }
            str = "";

            // Reverse
            while (cellName.Length > 1)
            {
                str += cellName.Substring((cellName.Length - 1), 1);
                cellName = cellName.Substring(0, cellName.Length - 1);
            }

            str += cellName;
            cellName = str;
            return cellName;

        }

        private string CellName(int nColumn, int nRow)
        {
            string cellName, str;

            cellName = ColumnName(nColumn);
            str = nRow.ToString();
            cellName += str;
            return cellName;
        }

    }
}
