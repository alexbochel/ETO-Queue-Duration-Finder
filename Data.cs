using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ETOCBurgDuration
{
    public class Data
    {
        public string MRPC;
        public string salesNum;
        public string mat;
        public string desc;
        public string dateCreated;
        public int daysAvailable;
        public int daysNotAvailable;

        /// <summary>
        /// Set initial data values. 
        /// </summary>
        public Data()
        {
            daysAvailable = 0;
            daysNotAvailable = 0;
        }

        /// <summary>
        /// Finally using something I learned in 2114!!!!!!!!
        /// </summary>
        /// <param name="otherData"> Data being compared to current data. </param>
        /// <returns> Whether or not the data is equal. </returns>
        public bool equals(Data otherData)
        {
            if (this.salesNum != otherData.salesNum)
            {
                return false;
            }
            else if (this.mat != otherData.mat)
            {
                return false;
            }
            else if (this.desc != otherData.desc)
            {
                return false;
            }
            else
            {
                return true; // Note: This should probably be unit tested. 
            }
        }
    }
}
