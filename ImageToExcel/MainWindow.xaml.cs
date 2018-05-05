using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Application = Microsoft.Office.Interop.Excel.Application;
using Window = System.Windows.Window;

namespace ImageToExcel
{
    /// <summary>
    ///     MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        ///     EXCEL主程序
        /// </summary>
        private Application excelApp;

        /// <summary>
        ///     工作簿
        /// </summary>
        private Workbooks workbooks;

        public MainWindow()
        {
            InitializeComponent();
        }

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        private async void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            // 清空之前的数据
            lblMessage.Content = string.Empty;
            lblUrl.Content = string.Empty;
            imgPreview.Source = null;
            // 打开图片
            var fileDialog = new OpenFileDialog { Filter = "图片文件 (*.jpg)|*.jpg" };
            fileDialog.Title = "请选择图片";
            fileDialog.ShowDialog();
            lblUrl.Content = fileDialog.FileName;

            if (string.IsNullOrEmpty(lblUrl.Content.ToString())) return;

            imgPreview.Source = ReadPicture(fileDialog.FileName);

            // 创建用于转换图片的线程
            btnBrowse.IsEnabled = false;
            gifLoading.StartAnimate();
            await ConvertAsync();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            CloseExcel();
        }

        /// <summary>
        ///     将图片转换成EXCEL
        /// </summary>
        private async Task ConvertAsync()
        {
            await Task.Run(() =>
            {
                string url = null;
                Dispatcher.Invoke(() =>
                {
                    lblMessage.Content = "转换开始";
                    url = lblUrl.Content.ToString();
                });
                var bmp = new Bitmap(url);
                try
                {
                    var sw = new Stopwatch();
                    sw.Start();
                    var width = bmp.Width;
                    var height = bmp.Height;
                    var excelPath = Path.Combine(Path.GetDirectoryName(url),
                        Path.GetFileNameWithoutExtension(url) + ".xlsx");

                    excelApp = new Application();
                    workbooks = excelApp.Workbooks;
                    // 创建工作簿
                    var workbook = workbooks.Add(true);
                    // 创建工作表
                    var worksheet = workbook.Worksheets[1] as Worksheet;

                    // Excel可视化
                    excelApp.Visible = true;

                    // 设置宽度和高度
                    worksheet.Cells.RowHeight = 7.5;
                    worksheet.Cells.ColumnWidth = 0.77;

                    for (var rowIndex = 0; rowIndex < height; rowIndex++)
                        for (var colIndex = 0; colIndex < width; colIndex++)
                        {
                            var color = bmp.GetPixel(colIndex, rowIndex);
                            var range = (Range)worksheet.Cells[rowIndex + 1, colIndex + 1];
                            // 需要将C#的颜色转成EXCEL的颜色
                            range.Interior.Color = Color.FromArgb(color.A, color.B, color.G, color.R).ToArgb();
                        }

                    workbook.SaveAs(excelPath);

                    sw.Stop();
                    var ts = sw.Elapsed;
                    Dispatcher.Invoke(() => { lblMessage.Content = $"转换完成,耗时{ts.Hours}时{ts.Minutes}分{ts.Seconds}秒"; });
                }
                catch
                {
                    Dispatcher.Invoke(() => { lblMessage.Content = "转换失败"; });
                }
                finally
                {
                    Dispatcher.Invoke(() =>
                    {
                        gifLoading.StopAnimate();
                        btnBrowse.IsEnabled = true;
                    });
                    // 释放图片占用资源
                    bmp.Dispose();
                    // 关闭EXCEL资源
                    CloseExcel();
                }
            });
        }

        /// <summary>
        ///     关闭EXCEL
        /// </summary>
        private void CloseExcel()
        {
            if (excelApp == null) return;
            try
            {
                // 关闭Excel进程
                var hwnd = new IntPtr(excelApp.Hwnd);
                GetWindowThreadProcessId(hwnd, out var processId);
                var process = Process.GetProcessById(processId);
                process.Kill();
            }
            catch
            {
                // ignored
            }
        }

        /// <summary>
        ///     读取图片并释放资源
        /// </summary>
        /// <param name="path">图片路径</param>
        /// <returns></returns>
        private BitmapImage ReadPicture(string path)
        {
            var bitmap = new BitmapImage();

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