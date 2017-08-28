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
    class Program
    {

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
