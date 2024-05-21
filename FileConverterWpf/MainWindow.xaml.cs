using OutlookApp = Microsoft.Office.Interop.Outlook.Application;
using ICSharpCode.SharpZipLib.GZip;
using ICSharpCode.SharpZipLib.Tar;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace FileConverterWpf
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "TGZ files (*.tgz)|*.tgz";
            if (openFileDialog.ShowDialog() == true)
            {
                FilePathTextBox.Text = openFileDialog.FileName;
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            string tgzFilePath = FilePathTextBox.Text;
            if (string.IsNullOrEmpty(tgzFilePath) || !File.Exists(tgzFilePath))
            {
                MessageBox.Show("Please select a valid .TGZ file.");
                return;
            }

            string outputDir = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(tgzFilePath));
            Directory.CreateDirectory(outputDir);

            try
            {
                ExtractTgz(tgzFilePath, outputDir);

                // Implementar lógica de conversão para .PST
                string pstFilePath = Path.Combine(outputDir, "output.pst");
                ConvertToPst(outputDir, pstFilePath);

                // Mostrar o caminho do arquivo convertido
                OutputTextBox.Text = $"File converted successfully: {pstFilePath}";
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void ExtractTgz(string tgzFilePath, string outputDir)
        {
            using (Stream inStream = File.OpenRead(tgzFilePath))
            using (Stream gzipStream = new GZipInputStream(inStream))
            {
                TarArchive tarArchive = TarArchive.CreateInputTarArchive(gzipStream);
                tarArchive.ExtractContents(outputDir);
                tarArchive.Close();
            }
        }

        private void ConvertToPst(string inputDir, string pstFilePath)
        {
            try
            {
                OutlookApp outlookApp = new OutlookApp();
                NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

                // Adiciona o arquivo PST ao perfil do Outlook
                outlookNamespace.AddStore(pstFilePath);
                MAPIFolder rootFolder = GetRootFolder(outlookNamespace, pstFilePath);

                // Exemplo: Adicionando um email de teste
                MailItem mailItem = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
                mailItem.Subject = "Test Email";
                mailItem.Body = "This is a test email.";
                mailItem.To = "example@example.com";
                mailItem.Save();

                // Salvando no PST
                rootFolder.Items.Add(mailItem);
            }
            catch (COMException comEx)
            {
                MessageBox.Show($"COM error: {comEx.Message}");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"General error: {ex.Message}");
            }
        }

        private MAPIFolder GetRootFolder(NameSpace outlookNamespace, string pstFilePath)
        {
            foreach (MAPIFolder folder in outlookNamespace.Folders)
            {
                if (folder.FullFolderPath.EndsWith(Path.GetFileNameWithoutExtension(pstFilePath)))
                {
                    return folder;
                }
            }
            throw new System.Exception("PST file root folder not found.");
        }
    }
}
