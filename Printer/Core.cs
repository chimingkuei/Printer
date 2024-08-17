using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Printer
{
    public class PrintStringAndImage
    {
        private PrintDocument printDocument;
        private string textToPrint;
        private Image imageToPrint;
        private int currentPage;

        private void PrintPageHandler(object sender, PrintPageEventArgs e)
        {
            switch (currentPage)
            {
                case 0:
                    // 設定要列印的字型和位置
                    Font font = new Font("Arial", 12);
                    float textX = 100; // 字串的X座標
                    float textY = 100; // 字串的Y座標
                    e.Graphics.DrawString(textToPrint, font, Brushes.Black, textX, textY);
                    break;

                case 1:
                    // 設定圖片的列印位置
                    float imageX = 100; // 圖片的X座標
                    float imageY = 150; // 圖片的Y座標，在字串下方
                    e.Graphics.DrawImage(imageToPrint, imageX, imageY, imageToPrint.Width, imageToPrint.Height);
                    break;
            }

            // 判斷是否還有下一頁
            currentPage++;
            e.HasMorePages = (currentPage < 2); // 總共兩頁
        }

        public PrintStringAndImage(string text, Image image)
        {
            textToPrint = text;
            imageToPrint = image;
            printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintPageHandler);
            printDocument.PrinterSettings.PrinterName = "Canon G4010 series"; // 如果不指定，會使用預設的印表機
            currentPage = 0;
            printDocument.DefaultPageSettings.PaperSize = new PaperSize("A4", 827, 1169);// 設定紙張大小為A4，單位inch
            printDocument.DefaultPageSettings.Landscape = false; // 設定為直向列印
        }

        public void Print()
        {
            try
            {
                printDocument.Print(); // 開始列印，不顯示對話框
            }
            catch (Exception ex)
            {
                Console.WriteLine("列印失敗: " + ex.Message);
            }
        }

        public void ShowPrintPreview()
        {
            PrintPreviewDialog previewDialog = new PrintPreviewDialog
            {
                Document = printDocument,
                Width = 800,
                Height = 600
            };

            // 顯示預覽對話框
            previewDialog.ShowDialog();
        }
    }
}
