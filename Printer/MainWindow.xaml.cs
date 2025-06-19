using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Net.Mime.MediaTypeNames;
using static Printer.BaseLogRecord;
using System.Windows.Media.Media3D;
using System.Drawing.Printing;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;


namespace Printer
{
    #region Config Class
    public class SerialNumber
    {
        [JsonProperty("Parameter1_val")]
        public string Parameter1_val { get; set; }
        [JsonProperty("Parameter2_val")]
        public string Parameter2_val { get; set; }
    }

    public class Model
    {
        [JsonProperty("SerialNumbers")]
        public SerialNumber SerialNumbers { get; set; }
    }

    public class RootObject
    {
        [JsonProperty("Models")]
        public List<Model> Models { get; set; }
    }
    #endregion

    public partial class MainWindow : System.Windows.Window
    {
        
        public MainWindow()
        {
            InitializeComponent();
        }

        #region Function
        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (MessageBox.Show("請問是否要關閉？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

        #region Config
        private SerialNumber SerialNumberClass()
        {
            SerialNumber serialnumber_ = new SerialNumber
            {
                Parameter1_val = Parameter1.Text,
                Parameter2_val = Parameter2.Text
            };
            return serialnumber_;
        }

        private void LoadConfig(int model, int serialnumber, bool encryption = false)
        {
            List<RootObject> Parameter_info = Config.Load(encryption);
            if (Parameter_info != null)
            {
                Parameter1.Text = Parameter_info[model].Models[serialnumber].SerialNumbers.Parameter1_val;
                Parameter2.Text = Parameter_info[model].Models[serialnumber].SerialNumbers.Parameter2_val;
            }
            else
            {
                // 結構:2個Models、Models下在各2個SerialNumbers
                SerialNumber serialnumber_ = SerialNumberClass();
                List<Model> models = new List<Model>
                {
                    new Model { SerialNumbers = serialnumber_ },
                    new Model { SerialNumbers = serialnumber_ }
                };
                List<RootObject> rootObjects = new List<RootObject>
                {
                    new RootObject { Models = models },
                    new RootObject { Models = models }
                };
                Config.SaveInit(rootObjects, encryption);
            }
        }
       
        private void SaveConfig(int model, int serialnumber, bool encryption = false)
        {
            Config.Save(model, serialnumber, SerialNumberClass(), encryption);
        }
        #endregion

        #region Dispatcher Invoke 
        public string DispatcherGetValue(TextBox control)
        {
            string content = "";
            this.Dispatcher.Invoke(() =>
            {
                content = control.Text;
            });
            return content;
        }

        public void DispatcherSetValue(string content, TextBox control)
        {
            this.Dispatcher.Invoke(() =>
            {
                control.Text = content;
            });
        }
        #endregion

        private bool IsPhysicalPrinter(string printerName)
        {
            // 篩選條件，可以根據虛擬印表機的名稱進行排除
            string[] virtualPrinters = { "OneNote", "PDF", "XPS", "FAX", "Microsoft Print to PDF", "AnyDesk" };
            foreach (string virtualPrinter in virtualPrinters)
            {
                if (printerName.Contains(virtualPrinter))
                {
                    return false;
                }
            }
            // 如果不包含虛擬印表機的名稱，則認為是實體印表機
            return true;
        }
        #endregion

        #region Parameter and Init
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadConfig(0, 0);
            Parameter1.Visibility = Visibility.Collapsed;
            Parameter2.Visibility = Visibility.Collapsed;
            Show_Printer.Visibility = Visibility.Collapsed;
            Print.Visibility = Visibility.Collapsed;
            PrintPreview.Visibility = Visibility.Collapsed;
        }
        BaseConfig<RootObject> Config = new BaseConfig<RootObject>();
        BaseLogRecord Logger = new BaseLogRecord();
        DataHandler DH = new DataHandler();
        #endregion

        #region Main Screen
        private void Main_Btn_Click(object sender, RoutedEventArgs e)
        {
            switch ((sender as Button).Name)
            {
                case nameof(Open_MergePdf_Folder):
                    {
                        System.Windows.Forms.FolderBrowserDialog pdfFolderPath = new System.Windows.Forms.FolderBrowserDialog();
                        pdfFolderPath.Description = "Choose Merge Pdf Folder Path";
                        pdfFolderPath.ShowDialog();
                        MergePdf_Folder_Path.Text = pdfFolderPath.SelectedPath;
                        break;
                    }
                case nameof(Open_SplitPdf_Folder):
                    {
                        System.Windows.Forms.FolderBrowserDialog pdfFolderPath = new System.Windows.Forms.FolderBrowserDialog();
                        pdfFolderPath.Description = "Choose Split Pdf Folder Path";
                        pdfFolderPath.ShowDialog();
                        SplitPdf_Folder_Path.Text = pdfFolderPath.SelectedPath;
                        break;
                    }
                case nameof(Show_Printer):
                    {
                        //PrintDocument printDoc = new PrintDocument();
                        //String sDefaultPrinter = printDoc.PrinterSettings.PrinterName;  // 取得預設的印表機名稱
                        foreach (string printer in PrinterSettings.InstalledPrinters)
                        {
                            if (IsPhysicalPrinter(printer))
                            {
                                Logger.WriteLog(printer, LogLevel.General, richTextBoxGeneral);
                            }
                        }
                        break;
                    }
                case nameof(Print):
                    {
                        string text = "這是一段要列印的文字內容。";
                        System.Drawing.Image image = System.Drawing.Image.FromFile(@"Icon\Rem Pin.bmp"); // 替換為圖片的路徑
                        PrinterHandler printer = new PrinterHandler(text, image);
                        printer.Print();
                        break;
                    }
                case nameof(PrintPreview):
                    {
                        string text = "這是一段要列印的文字內容。";
                        System.Drawing.Image image = System.Drawing.Image.FromFile(@"Icon\Rem Pin.bmp"); // 替換為圖片的路徑
                        PrinterHandler printer = new PrinterHandler(text, image);
                        printer.ShowPrintPreview();
                        break;
                    }
                case nameof(Merge_PDF):
                    {
                        string[] pdfFiles = new string[]
                        {
                            @"D:\Chimingkuei\repos\Printer\Testing Set\File_20250614_0001.pdf",
                            @"D:\Chimingkuei\repos\Printer\Testing Set\File_20250614_0002.pdf",
                            @"D:\Chimingkuei\repos\Printer\Testing Set\File_20250614_0003.pdf",
                            @"D:\Chimingkuei\repos\Printer\Testing Set\File_20250614_0004.pdf",
                            @"D:\Chimingkuei\repos\Printer\Testing Set\File_20250614_0005.pdf",
                        };
                        string outputFilePath = @"D:\Chimingkuei\repos\Printer\MergedOutput.pdf";
                        DH.MergePdfFiles(pdfFiles, outputFilePath, 0);
                        Console.WriteLine("PDF 合併完成！");
                        break;
                    }
                case nameof(Split_PDF):
                    {
                        string inputFile = @"C:\PDFs\input.pdf";
                        string outputFolder = @"C:\PDFs\Split\";
                        DH.SplitPdfFiles(inputFile, outputFolder);
                        Console.WriteLine("PDF 分割完成！");
                        break;
                    }
            }
        }
        #endregion
        

    }
}
