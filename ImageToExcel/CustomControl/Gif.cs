using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace ImageToExcel.CustomControl
{
    public class Gif : System.Windows.Controls.Image
    {
        /// <summary>  
        /// gif动画的System.Drawing.Bitmap  
        /// </summary>  
        private Bitmap gifBitmap;

        /// <summary>  
        /// 用于显示每一帧的BitmapSource  
        /// </summary>  
        private BitmapSource bitmapSource;

        /// <summary>
        /// 图片路径
        /// </summary>
        public string GifPath { get; set; }

        /// <summary>
        /// 初期化
        /// </summary>
        private void Init()
        {
            // 从程序集资源中获取 
            Stream stream = Application.GetResourceStream(new Uri(GifPath, UriKind.Absolute)).Stream;
            this.gifBitmap = new Bitmap(stream);
            this.bitmapSource = this.GetBitmapSource();
            this.Source = this.bitmapSource;
        }

        /// <summary>  
        /// 从System.Drawing.Bitmap中获得用于显示的那一帧图像的BitmapSource  
        /// </summary>  
        /// <returns></returns>  
        private BitmapSource GetBitmapSource()
        {
            IntPtr handle = IntPtr.Zero;

            try
            {
                handle = this.gifBitmap.GetHbitmap();
                this.bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(handle, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {
                if (handle != IntPtr.Zero)
                {
                    DeleteObject(handle);
                }
            }

            return this.bitmapSource;
        }

        /// <summary>  
        /// Start animation  
        /// </summary>  
        public void StartAnimate()
        {
            if (this.gifBitmap == null)
            {
                // 尚未初期化
                this.Init();
            }
            this.Visibility = Visibility.Visible;
            ImageAnimator.Animate(this.gifBitmap, this.OnFrameChanged);
        }

        /// <summary>  
        /// Stop animation  
        /// </summary>  
        public void StopAnimate()
        {
            ImageAnimator.StopAnimate(this.gifBitmap, this.OnFrameChanged);
            this.Visibility = Visibility.Collapsed;
        }

        /// <summary>  
        /// Event handler for the frame changed  
        /// </summary>  
        private void OnFrameChanged(object sender, EventArgs e)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                ImageAnimator.UpdateFrames();  // 更新到下一帧  
                if (this.bitmapSource != null)
                {
                    this.bitmapSource.Freeze();
                }

                //// Convert the bitmap to BitmapSource that can be display in WPF Visual Tree  
                this.bitmapSource = this.GetBitmapSource();
                Source = this.bitmapSource;
                this.InvalidateVisual();
            }));
        }

        /// <summary>  
        /// Delete local bitmap resource  
        /// Reference: http://msdn.microsoft.com/en-us/library/dd183539(VS.85).aspx  
        /// </summary>  
        [DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool DeleteObject(IntPtr hObject);
    }

}
