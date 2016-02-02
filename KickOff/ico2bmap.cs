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
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        [DllImport("Shell32.dll", EntryPoint = "DestroyIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int DestroyIconEx(IntPtr pIcon);

        static private IconBitMap ExtractIconBitMap(Icon ico)
        {
            IconBitMap ibm = null;

            try
            {
                Bitmap bm = ico.ToBitmap();

                BitmapSource bmsource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    bm.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

                ibm = new IconBitMap()
                {
                    BitmapSize = bm.Width,
                    //bitmap = bm,
                    bitmapsource = bmsource
                    //writeablebitmap = wbitmap
                };
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            return ibm;
        }

        static private IconBitMap ExtractIconBitMap(IntPtr pi)
        {
            return ExtractIconBitMap(Icon.FromHandle(pi));
        }
        private static int ExtractIconEx_mine(string file, int iconindex, IntPtr ico_l, IntPtr ico_s)
        {
            return ExtractIconEx(file, iconindex, out ico_l, out ico_s, 1); ;
        }

        static IconBitMap ExtractIconBitMapFromFile(string file, int iconindex, bool largeIcon)
        {
            IntPtr large = IntPtr.Zero;
            IntPtr small = IntPtr.Zero;
            IconBitMap ibm = null;
            try
            {
                int i = ExtractIconEx_mine(file, -1, IntPtr.Zero, IntPtr.Zero);

                if (i > iconindex)
                {
                    ExtractIconEx(file, iconindex, out large, out small, 1);
                    if ((largeIcon && IntPtr.Zero == large ) || IntPtr.Zero == small)
                    {
                        //throw new Exception("Did not get valid icons");
                        ibm = ExtractIconBitMap(SystemIcons.Error);
                    }
                    else
                    {
                        ibm = ExtractIconBitMap(largeIcon ? large : small);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            //finally
            //{
            //    // destroy the icon from Win32 or memory will not be released
            //    if(IntPtr.Zero != large)
            //        DestroyIconEx(large);
            //    if (IntPtr.Zero != small)
            //        DestroyIconEx(small);
            //}
            return ibm;
        }

        public static IconBitMap GetBitmapFromFileIcon(string file)
        {
            IconBitMap ibm = null;
            //System.Drawing.Icon ico = SystemIcons.Error; //.WinLogo;

            // if a directory path is used just get the system icon
            if (File.GetAttributes(file).HasFlag(FileAttributes.Directory))
            {
                // get a 'built in' icon for a folder
                try
                {
                    ibm = ExtractIconBitMapFromFile("shell32.dll", 4, true);
                    //ico = SystemIcons.Hand; //System.Drawing.Icon.ExtractAssociatedIcon(% SystemRoot %\system32\shell32.dll); C:\Windows\System32\imageres.dll % SystemRoot %\system32\DDORes.dll
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message);
                }

            }
            else
            {
                try
                {
                    ibm = ExtractIconBitMapFromFile(file, 0, true);
                    if (null == ibm)
                    {
                        System.Drawing.Icon ico = System.Drawing.Icon.ExtractAssociatedIcon(file);
                        ibm = ExtractIconBitMap(ico);
                    }
                }
                catch //(Exception e)
                {
                    //throw new Exception(e.Message + "File: " + file);
                    if (null == ibm)
                    {
                        System.Drawing.Icon ico = System.Drawing.Icon.ExtractAssociatedIcon(file);
                        ibm = ExtractIconBitMap(ico);
                    }
                }
            }

            // Convert the icon to bitmap
            /*
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
*/

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
