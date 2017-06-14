using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Diagnostics;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;

namespace ImageToExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        /// <summary>
        /// EXCEL主程序
        /// </summary>
        Microsoft.Office.Interop.Excel.Application excelApp;

        /// <summary>
        /// 工作簿
        /// </summary>
        Workbooks workbooks;

        /// <summary>
        /// 用于转换图片的线程
        /// </summary>
        Thread th;

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public MainWindow()
        {
            InitializeComponent();

        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            // 清空之前的数据
            this.lblMessage.Content = string.Empty;
            this.lblUrl.Content = string.Empty;
            this.imgPreview.Source = null;
            // 打开图片
            OpenFileDialog fileDialog = new OpenFileDialog() { Filter = "图片文件 (*.jpg)|*.jpg" };
            fileDialog.Title = "请选择图片";
            fileDialog.ShowDialog();
            this.lblUrl.Content = fileDialog.FileName;

            if (string.IsNullOrEmpty(this.lblUrl.Content.ToString()))
            {
                return;
            }

            this.imgPreview.Source = this.ReadPicture(fileDialog.FileName);

            // 创建用于转换图片的线程
            this.th = new Thread(Convert);
            this.btnBrowse.IsEnabled = false;
            this.gifLoading.StartAnimate();
            this.th.Start();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            this.th?.Abort();
            this.CloseExcel();
        }

        /// <summary>
        /// 将图片转换成EXCEL
        /// </summary>
        private void Convert()
        {
            string url = null;
            Dispatcher.Invoke(() =>
            {
                this.lblMessage.Content = "转换开始";
                url = this.lblUrl.Content.ToString();
            });
            Bitmap bmp = new Bitmap(url);
            try
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();
                int width = bmp.Width;
                int height = bmp.Height;
                string excelPath = $@"{Path.GetDirectoryName(url)}\{Path.GetFileNameWithoutExtension(url)}.xlsx";

                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbooks = excelApp.Workbooks;
                // 创建工作簿
                Workbook workbook = workbooks.Add(true);
                // 创建工作表
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;

                // Excel可视化
                excelApp.Visible = true;

                // 设置宽度和高度
                worksheet.Cells.RowHeight = 7.5;
                worksheet.Cells.ColumnWidth = 0.77;

                for (int rowIndex = 0; rowIndex < height; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < width; colIndex++)
                    {
                        Color color = bmp.GetPixel(colIndex, rowIndex);
                        Range range = (Range)worksheet.Cells[rowIndex + 1, colIndex + 1];
                        // 需要将C#的颜色转成EXCEL的颜色
                        range.Interior.Color = Color.FromArgb(color.A, color.B, color.G, color.R).ToArgb();
                    }
                }

                workbook.SaveAs(excelPath);

                sw.Stop();
                TimeSpan ts = sw.Elapsed;
                Dispatcher.Invoke(() =>
                {
                    this.lblMessage.Content = $"转换完成,耗时{ts.Hours}时{ts.Minutes}分{ts.Seconds}秒";
                });
            }
            catch
            {
                Dispatcher.Invoke(() =>
                {
                    this.lblMessage.Content = "转换失败";
                });
            }
            finally
            {
                Dispatcher.Invoke(() =>
                {
                    this.gifLoading.StopAnimate();
                    this.btnBrowse.IsEnabled = true;
                });
                // 释放图片占用资源
                bmp.Dispose();
                // 关闭EXCEL资源
                this.CloseExcel();
                // 关闭线程
                this.th.Abort();
            }
        }

        /// <summary>
        /// 关闭EXCEL
        /// </summary>
        private void CloseExcel()
        {
            if (excelApp != null)
            {
                try
                {
                    // 关闭Excel进程
                    IntPtr hwnd = new IntPtr(excelApp.Hwnd);
                    int processId = 0;
                    GetWindowThreadProcessId(hwnd, out processId);
                    System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(processId);
                    process.Kill();
                }
                catch
                {

                }
            }

        }

        /// <summary>
        /// 读取图片并释放资源
        /// </summary>
        /// <param name="path">图片路径</param>
        /// <returns></returns>
        private BitmapImage ReadPicture(string path)
        {
            BitmapImage bitmap = new BitmapImage();

            if (File.Exists(path))
            {
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;

                using (Stream ms = new MemoryStream(File.ReadAllBytes(path)))
                {
                    bitmap.StreamSource = ms;
                    bitmap.EndInit();
                    bitmap.Freeze();
                }
            }

            return bitmap;
        }
    }
}
