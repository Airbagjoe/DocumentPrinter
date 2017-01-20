using DocumentPrinter.Printers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentPrinter.Printers
{
    public class PDFPrinter : BasePrinter
    {
        public override bool Print(string pdfFileName, string printerName)
        {
            try
            {
                if (!base.Print(pdfFileName, printerName))
                {
                    return false;
                }

                KillAdobe("AcroRd32");

                return true;
            }
            catch
            {
                ShowPrintError(pdfFileName);
                return false;
            }
        }

        //For whatever reason, sometimes adobe likes to be a stage 5 clinger.
        //So here we kill it with fire.
        private static bool KillAdobe(string name)
        {
            foreach (Process clsProcess in Process.GetProcesses().Where(
                         clsProcess => clsProcess.ProcessName.StartsWith(name)))
            {
                clsProcess.Kill();
                return true;
            }
            return false;
        }
    }//END Class}
}
