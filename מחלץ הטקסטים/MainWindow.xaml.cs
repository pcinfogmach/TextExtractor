using Docnet.Core;
using Docnet.Core.Models;
using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;


namespace מחלץ_הטקסטים
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

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                progressBar.IsIndeterminate = true;
                try
                {
                    string fileName = openFileDialog.FileName;
                    string directory = Path.GetDirectoryName(fileName);
                    string newTextFileName = Path.Combine(directory, Path.GetFileNameWithoutExtension(fileName) + ".txt");
                    string content = await Task.Run(() => GetFileContent(fileName));
                    File.WriteAllText(newTextFileName, content);
                    System.Diagnostics.Process.Start(newTextFileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                progressBar.IsIndeterminate = false;
            }
        }

        string GetFileContent(string filePath)
        {
            try
            {
                filePath = filePath.ToLower();
                if (filePath.EndsWith(".pdf"))
                {
                    return DocNetPdfTextExtractor(filePath);
                }
                else
                {
                    return Toxy.ParserFactory.CreateText(new Toxy.ParserContext(filePath)).Parse();
                }
            }
            catch (System.NotSupportedException)
            {
                var textExtractor = new TikaOnDotNet.TextExtraction.TextExtractor();
                var result = textExtractor.Extract(filePath);
                return result.Text;
            }
            catch
            {
                return string.Empty;
            }
        }


        string DocNetPdfTextExtractor(string filePath)
        {
            var pageTextDictionary = new ConcurrentDictionary<int, string>();
            using (var docReader = DocLib.Instance.GetDocReader(filePath, new PageDimensions()))
            {
                int pageCount = docReader.GetPageCount();
                Docnet.Core.Readers.IPageReader pageReader = null;
                Parallel.For(0, pageCount, i =>
                {
                    pageReader = docReader.GetPageReader(i);
                    pageTextDictionary[i] = pageReader.GetText();
                });
                pageReader.Dispose();
            }

            StringBuilder stb = new StringBuilder();

            for (int i = 0; i < pageTextDictionary.Count; i++)
            {
                stb.AppendLine(pageTextDictionary[i]);
            }

            return stb.ToString();
        }
    }
}