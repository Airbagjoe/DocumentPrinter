using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentPrinter
{
    public class PDFPrinter
    {
        public static bool PrintPDF(string pdfFileName)
        {
            try
            {
                var p = new Process();
                p.StartInfo = new ProcessStartInfo()
                {
                    CreateNoWindow = true,
                    Verb = "print",
                    FileName = pdfFileName //put the correct path here
                };
                p.Start();
                /*
                Process proc = new Process();
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc.StartInfo.Verb = "print";

                //Define location of adobe reader/command line
                //switches to launch adobe in "print" mode
                proc.StartInfo.FileName =
                  @"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe";
                proc.StartInfo.Arguments = String.Format(@"/p /h {0}", pdfFileName);
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.CreateNoWindow = true;

                proc.Start();
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                if (proc.HasExited == false)
                {
                    proc.WaitForExit(10000);
                }

                proc.EnableRaisingEvents = true;
                  * */

                if (!p.HasExited)
                {
                    p.WaitForExit(10000);
                }
                p.Close();
                KillAdobe("AcroRd32");
                return true;
            }
            catch
            {
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
