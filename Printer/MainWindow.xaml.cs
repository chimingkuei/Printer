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
using Microsoft.Win32;


namespace Printer
{
    #region Config Class
    public class SerialNumber
    {
        [JsonProperty("MergePdf_Folder_Path_val")]
        public string MergePdf_Folder_Path_val { get; set; }
        [JsonProperty("SplitPdf_Folder_Path_val")]
        public string SplitPdf_Folder_Path_val { get; set; }
        [JsonProperty("RotatePdf_Folder_Path_val")]
        public string RotatePdf_Folder_Path_val { get; set; }
        [JsonProperty("Angle_val")]
        public string Angle_val { get; set; }
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
                MergePdf_Folder_Path_val = MergePdf_Folder_Path.Text,
                SplitPdf_Folder_Path_val = SplitPdf_Folder_Path.Text,
                RotatePdf_Folder_Path_val = RotatePdf_Folder_Path.Text,
                Angle_val = Angle.Text
            };
            return serialnumber_;
        }

        private void LoadConfig(int model, int serialnumber, bool encryption = false)
        {
            List<RootObject> Parameter_info = Config.Load(encryption);
            if (Parameter_info != null)
            {
                MergePdf_Folder_Path.Text = Parameter_info[model].Models[serialnumber].SerialNumbers.MergePdf_Folder_Path_val;
                SplitPdf_Folder_Path.Text = Parameter_info[model].Models[serialnumber].SerialNumbers.SplitPdf_Folder_Path_val;
                RotatePdf_Folder_Path.Text = Parameter_info[model].Models[serialnumber].SerialNumbers.RotatePdf_Folder_Path_val;
                Angle.Text = Parameter_info[model].Models[serialnumber].SerialNumbers.Angle_val;
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
                        Logger.WriteLog("開啟合併PDF資料夾!", LogLevel.General, richTextBoxGeneral);
                        break;
                    }
                case nameof(Open_SplitPdf_File):
                    {
                        OpenFileDialog openFileDialog = new OpenFileDialog();
                        openFileDialog.Filter = "Files|*.pdf;*|All files|*.*";
                        if (openFileDialog.ShowDialog() == true)
                        {
                            SplitPdf_Folder_Path.Text = openFileDialog.FileName;
                        }
                        Logger.WriteLog("開啟切割PDF檔!", LogLevel.General, richTextBoxGeneral);
                        break;
                    }
                case nameof(Open_RotatePdf_File):
                    {
                        OpenFileDialog openFileDialog = new OpenFileDialog();
                        openFileDialog.Filter = "Files|*.pdf;*|All files|*.*";
                        if (openFileDialog.ShowDialog() == true)
                        {
                            RotatePdf_Folder_Path.Text = openFileDialog.FileName;
                        }
                        Logger.WriteLog("開啟旋轉PDF檔!", LogLevel.General, richTextBoxGeneral);
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
                        if (!string.IsNullOrEmpty(MergePdf_Folder_Path.Text))
                        {
                            var sortedPdfFiles = Directory.GetFiles(MergePdf_Folder_Path.Text, "*.pdf")
                            .OrderBy(path =>
                            {
                                string fileName = System.IO.Path.GetFileNameWithoutExtension(path);
                                return int.Parse(fileName);
                            })
                            .ToArray();
                            // 顯示排序檔名
                            //foreach (var file in sortedPdfFiles)
                            //{
                            //    Console.WriteLine(file);
                            //}
                            Directory.CreateDirectory(System.IO.Path.Combine(MergePdf_Folder_Path.Text, "MergePdf"));
                            string outputFilePath = System.IO.Path.Combine(MergePdf_Folder_Path.Text, "MergePdf", "MergeResult.pdf");
                            string selectedText = Angle.Text;
                            int angle = Convert.ToInt32(selectedText.Replace("度", ""));
                            DH.MergeOrRotatePdfFiles(sortedPdfFiles, outputFilePath, angle, true);
                            Logger.WriteLog("合併PDF完成!", LogLevel.General, richTextBoxGeneral);
                        }
                        else
                        {
                            MessageBox.Show("請輸入欲合併PDF檔資料夾路徑!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        break;
                    }
                case nameof(Split_PDF):
                    {
                        if (!string.IsNullOrEmpty(SplitPdf_Folder_Path.Text))
                        {
                            string outputFolder = System.IO.Path.GetDirectoryName(SplitPdf_Folder_Path.Text);
                            DH.SplitPdfFiles(SplitPdf_Folder_Path.Text, outputFolder);
                            Logger.WriteLog("分割PDF完成!", LogLevel.General, richTextBoxGeneral);
                        }
                        else
                        {
                            MessageBox.Show("請輸入欲分割PDF檔資料夾路徑!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        break;
                    }
                case nameof(Rotate_PDF):
                    {
                        if (!string.IsNullOrEmpty(RotatePdf_Folder_Path.Text))
                        {
                            string outputFolder = System.IO.Path.GetDirectoryName(RotatePdf_Folder_Path.Text);
                            Directory.CreateDirectory(System.IO.Path.Combine(outputFolder, "RotatePdf"));
                            string outputFilePath = System.IO.Path.Combine(outputFolder, "RotatePdf", "RotateResult.pdf");
                            string[] sortedPdfFiles = new string[1];
                            sortedPdfFiles[0] = RotatePdf_Folder_Path.Text;
                            string selectedText = Angle.Text;
                            int angle = Convert.ToInt32(selectedText.Replace("度", ""));
                            DH.MergeOrRotatePdfFiles(sortedPdfFiles, outputFilePath, angle, false);
                            Logger.WriteLog("旋轉PDF完成!", LogLevel.General, richTextBoxGeneral);
                        }
                        else
                        {
                            MessageBox.Show("請輸入欲旋轉PDF檔資料夾路徑!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        break;
                    }
                case nameof(Save_Config):
                    {
                        SaveConfig(0, 0);
                        Logger.WriteLog("儲存參數!", LogLevel.General, richTextBoxGeneral);
                        break;
                    }
            }
        }
        #endregion
        

    }
}
