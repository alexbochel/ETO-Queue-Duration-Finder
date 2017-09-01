using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows;
using System.IO;

namespace ETOCBurgDuration
{
    /// <summary>
    /// Entry point. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 8/30/2017
    /// 
    /// </summary>
    class Program
    {
        /// <summary>
        /// Did you hear about the kidnapping at school? It's fine, he woke up.
        /// </summary>
        /// <param name="args"> Array of excel file paths. </param>
        [STAThread]
        static void Main(string[] args)
        {
            string[] files = null;
            
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    files = Directory.GetFiles(fbd.SelectedPath);

                    Console.WriteLine("Files found: " + files.Length.ToString(), "Message");
                }
            }

            DataLoader loader = new DataLoader(files);
            loader.execute();

        }

    }
}
