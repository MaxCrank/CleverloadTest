using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using Microsoft.Office.Interop.Word;

namespace CleverloadOpenXMLTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSelectFile_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Filter = "Word (*.doc;*.docx)|*.doc;*.docx";
            if (openFileDialog.ShowDialog() == true && openFileDialog.FileName.Length > 0)
            {
                lblFilePath.Text = openFileDialog.FileName;
                if (string.IsNullOrEmpty(openFileDialog.FileName) || !File.Exists(openFileDialog.FileName))
                {
                    btnTransform.IsEnabled = false;
                    MessageBox.Show("Please, provide a valid path to Microsoft Word document", "File does not exist",
                        MessageBoxButton.OK, MessageBoxImage.Stop);
                }
                else
                {
                    btnTransform.IsEnabled = true;
                    OpenXmlWorker.Instance.ConvertWordDocument(openFileDialog.FileName, WdSaveFormat.wdFormatDocumentDefault);
                    LoadWordToViewer();
                }
            }
        }

        private void BtnTransform_OnClick(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(Properties.Resources.TransformMessage, "Following actions will be performed",
                MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    OpenXmlWorker.Instance.Transform(tbFirstName.Text, tbLastName.Text, tbCity.Text);
                    if (!string.IsNullOrEmpty(OpenXmlWorker.Instance.LastConvertedWordFileName))
                        OpenXmlWorker.Instance.ConvertWordDocument(OpenXmlWorker.Instance.LastConvertedWordFileName, WdSaveFormat.wdFormatDocument97);
                    LoadWordToViewer();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Transformation failed. Probably environment is not set properly. Details: " + ex.Message);
                    OpenXmlWorker.Instance.RemoveLastDocxFile();
                    OpenXmlWorker.Instance.RemoveLastXpsFile();
                    xpsViewer.Document = null;
                }
            }
        }

        private void LoadWordToViewer()
        {
            OpenXmlWorker.Instance.ConvertWordDocument(OpenXmlWorker.Instance.LastConvertedWordFileName ??
                OpenXmlWorker.Instance.LastWordFileName, WdSaveFormat.wdFormatXPS);
            xpsViewer.Document = OpenXmlWorker.Instance.XpsDocument.GetFixedDocumentSequence();
        }

        protected override void OnClosed(EventArgs e)
        {
            OpenXmlWorker.Instance.RemoveLastDocxFile();
            OpenXmlWorker.Instance.RemoveLastXpsFile();
            xpsViewer.Document = null;
            base.OnClosed(e);
        }
    }
}
