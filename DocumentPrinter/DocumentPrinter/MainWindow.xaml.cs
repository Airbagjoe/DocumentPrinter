using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using DocumentPrinter.Properties;
using MessageBox = System.Windows.Forms.MessageBox;
using DocumentPrinter.Printers;

namespace DocumentPrinter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        private void bFolder_Click(object sender, RoutedEventArgs e)
        {
            var folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (!Directory.Exists(folderBrowser.SelectedPath))
                {
                    MessageBox.Show(@"Folder does not exist");
                    return;
                }

                tbFolder.Text = folderBrowser.SelectedPath;

            }
        }

        private async void bPrint_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbFolder.Text))
            {
                MessageBox.Show(@"No folder selected");
                return;
            }

            if (!Directory.Exists(tbFolder.Text))
            {
                MessageBox.Show(@"Folder does not exist");
                return;
            }

            bPrint.IsEnabled = false;
            bFolder.IsEnabled = false;

            var progress = new Progress<int>();
            progress.ProgressChanged += (o, i) =>
            {
                pBar.Value = i;
            };

            await Print(tbFolder.Text, progress);

            bPrint.IsEnabled = true;
            bFolder.IsEnabled = true;
        }


        private Task Print(string folder, IProgress<int> progres = null)
        {
            var print = false;
            var printerName = string.Empty;
            using (var pDialog = new System.Windows.Forms.PrintDialog())
            {
                printerName = pDialog.PrinterSettings.PrinterName;
                print = pDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK;
            }

            if (print != true)
            {
                return Task.FromResult(false);
            }

            return Task.Run(() =>
            {
                var directory = new DirectoryInfo(folder);
                var files = directory.GetFiles();
                var filtered = files.Where(f => (f.Attributes & FileAttributes.Hidden) == 0).ToList();
                decimal i = 1;
                foreach (var file in filtered)
                {
                    BasePrinter printer = null;

                    if (Settings.Default.AdobeExtensions.Contains(file.Extension, StringComparison.InvariantCultureIgnoreCase))
                    {
                        printer = new PDFPrinter();
                    }
                    else if (Settings.Default.WordExtensions.Contains(file.Extension, StringComparison.InvariantCultureIgnoreCase))
                    {
                        printer = new WordPrinter();
                    }

                    if (printer != null)
                    {
                        printer.Print(file.FullName, printerName);
                    }

                    if (progres != null)
                    {
                        progres.Report((int)((i / filtered.Count) * 100));
                    }

                    i++;
                }
            });
        }
    }
}
