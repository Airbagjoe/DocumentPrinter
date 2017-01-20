using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DocumentPrinter.Printers
{
    public abstract class BasePrinter
    {
        public virtual bool Print(string fileName, string printerName)
        {
            try
            {
                var p = new Process();
                p.StartInfo = new ProcessStartInfo()
                {
                    Arguments = "\"" + printerName + "\"",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true,
                    Verb = "PrintTo",
                    FileName = fileName //put the correct path here
                };

                p.Start();

                if (!p.HasExited)
                {
                    p.WaitForExit(10000);
                }
                p.Close();
                
                return true;
            }
            catch
            {
                ShowPrintError(fileName);
                return false;
            }
        }

        protected virtual void ShowPrintError(string fileName)
        {
            MessageBox.Show($"Failed to print {fileName}.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
