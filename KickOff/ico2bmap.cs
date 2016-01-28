using System;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Windows;
using System.IO;
using System.Runtime.InteropServices;


namespace KickOff
{
    public static class ico2bmap
    {
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        static Icon ExtractIcon(string file, int number, bool largeIcon)
        {
            IntPtr large;
            IntPtr small;
            ExtractIconEx(file, number, out large, out small, 1);
            try
            {
                return Icon.FromHandle(largeIcon ? large : small);
            }
            catch
            {
                return null;
            }

        }

        public static IconBitMap GetBitmapFromFileIcon(string file)
        {
            IconBitMap ibm = null;
            System.Drawing.Icon ico = SystemIcons.Error; //.WinLogo;

            // if a directory path is used just get the system icon
            if (File.GetAttributes(file).HasFlag(FileAttributes.Directory))
            {
                // get a 'built in' icon for a folder
                try {
                    ico = ExtractIcon("shell32.dll", 4, true);
                    //ico = SystemIcons.Hand; //System.Drawing.Icon.ExtractAssociatedIcon(% SystemRoot %\system32\shell32.dll); C:\Windows\System32\imageres.dll % SystemRoot %\system32\DDORes.dll
                }
                catch
                {
                    throw new Exception(string.Format("Extraction of icon failed: {0}", "System Icons"));
                }

            }
            else
            {
                try {
                    ico = System.Drawing.Icon.ExtractAssociatedIcon(file);
                }
                catch
                {
                    throw new Exception("Extraction of icon failed: " + file);
                }
            }
            /// Convert the icon to bitmap
            try {
                Bitmap bm = ico.ToBitmap();

                BitmapSource bmsource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    bm.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

                //WriteableBitmap wbitmap = new WriteableBitmap(bmsource);

                ibm = new IconBitMap()
                {
                    BitmapSize = bm.Width,
                    //bitmap = bm,
                    bitmapsource = bmsource
                    //writeablebitmap = wbitmap
                };

            }
            catch
            {
                    throw new Exception("Extraction of icon failed: " + file );
            }

            return ibm;
        }
    }

    public class IconBitMap
    {
        public IconBitMap()
        {
        }
        public int BitmapSize { get; set; }
        public BitmapSource bitmapsource { get; set; }

        //public Bitmap bitmap { get; set; }
        //public WriteableBitmap writeablebitmap { get; set; }
    }
}
