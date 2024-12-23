using System;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Windows;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;

namespace KickOff
{
    public static class ico2bmap
    {
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconEx", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int ExtractIconEx(string sFile, int iIndex, IntPtr[] piLargeVersion, IntPtr[] piSmallVersion, int amountIcons);

        [DllImport("user32.dll", EntryPoint = "DestroyIcon", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int DestroyIcon(IntPtr hIcon);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);

        public static int USE_MAIN_ICON = -1;

        public static IconBitMap ExtractIconBitMap(Icon ico)
        {
            IconBitMap ibm = null;

            Bitmap bm = null;
            IntPtr hbm = IntPtr.Zero;
            try
            {
                bm = ico.ToBitmap();
                hbm = bm.GetHbitmap();
                BitmapSource bmsource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    hbm, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

                ibm = new IconBitMap()
                {
                    BitmapSize = bm.Width,
                    bitmap = bm,
                    bitmapsource = bmsource
                };
            }
            catch (Exception e)
            {
                lnkio.WriteProgramLog("ExtractIconBitMap() failed: " + e.Message);
            }
            finally
            {
                //release all resources used
                if( IntPtr.Zero != hbm)
                    DeleteObject(hbm);

                if ( null != bm )
                    bm.Dispose();

                if ( null == ibm )
                    ibm = ExtractIconBitMap(SystemIcons.Error);
            }
            // must always return a valid icon bitmap 
            return ibm;
        }
        static private IconBitMap ExtractIconBitMap(IntPtr pi)
        {
            return ExtractIconBitMap(Icon.FromHandle(pi));
        }
        private static int ExtractIconEx_GetCount(string file, IntPtr[] ico_l, IntPtr[] ico_s)
        {
            return ExtractIconEx(file, -1, ico_l, ico_s, 1); ;
        }

        private static IconBitMap ExtractIconBitMapFromFile(string file, int iconindex)
        {
            IntPtr[] large = new IntPtr[1] { IntPtr.Zero };
            IntPtr[] small = new IntPtr[1] { IntPtr.Zero };
            IconBitMap ibm = null;
            try
            {
                int totalIcons = ExtractIconEx_GetCount(file, large, small);

                if (totalIcons > iconindex)
                {
                    ExtractIconEx(file, iconindex, large, small, 1);
                    if (large[0] != IntPtr.Zero)
                    {
                        ibm = ExtractIconBitMap(large[0]);
                        DestroyIcon(large[0]);
                        if (small[0] != IntPtr.Zero)
                            DestroyIcon(small[0]);
                    }
                }
            }
            catch (Exception e)
            {
                lnkio.WriteProgramLog("ExtractIconBitMapFromFile(): Failed to load icon(s) from " + 
                    file + " System Error: " + e.Message);
            }
            finally
            {
                try {
                    if (null == ibm)
                        ibm = ExtractIconBitMap(System.Drawing.Icon.ExtractAssociatedIcon(file));
                }
                catch (Exception e)
                {
                    lnkio.WriteProgramLog("ExtractIconBitMapFromFile(): Failed to load icon(s) from " +
                        file + " System Error: " + e.Message);
                }
                finally
                {
                    if (null == ibm)
                        ibm = ExtractIconBitMap(SystemIcons.Error);
                }
            }
            return ibm;
        }

        public const int MAX_ICONS = 500;

        public static ObservableCollection<IconBitMap> ExtractAllIconBitMapFromFile(string file)
        { return ExtractAllIconBitMapFromFile(file, MAX_ICONS); }

        private static ObservableCollection<IconBitMap> ExtractAllIconBitMapFromFile(string file, int maxIcons = 0)
        {
            // limit the maximum number to be retrieved
            if (maxIcons > MAX_ICONS) maxIcons = MAX_ICONS;

            IntPtr[] large = new IntPtr[maxIcons];
            IntPtr[] small = new IntPtr[maxIcons];
            IconBitMap iconbitmap = null;
            bool bReportError = false;
            ObservableCollection<IconBitMap> ibm = new ObservableCollection<IconBitMap>();

            try
            {
                if (maxIcons == USE_MAIN_ICON)
                {
                    iconbitmap = ExtractIconBitMap(System.Drawing.Icon.ExtractAssociatedIcon(file));
                    if (null != iconbitmap)
                        ibm.Add(iconbitmap);
                }
                else {
                    // The number of icons retrieved may or may not be the icon count total
                    // It my be all of the large and small combined.
                    int totalicons = ExtractIconEx_GetCount(file, large, small);
                    // totalicons will be the number of distinct icons large and small versions
                    if (totalicons > 0)
                    {
                        if (totalicons > maxIcons)
                            totalicons = maxIcons;
                        // iconsExtracted may be a count of all the large and small combined
                        // if the icon source is an EXE and be exactly double the query 
                        // number given above. 
                        int iconsExtracted = ExtractIconEx(file, 0, large, small, totalicons); // note arrays are passed as pointers, so no 'out'
                        if (iconsExtracted > 0 )
                        {
                            for (int i = 0; i < totalicons; i++)
                            {
                                if (IntPtr.Zero != large[i])
                                {
                                    iconbitmap = ExtractIconBitMap(large[i]);
                                    if (null != iconbitmap)
                                        ibm.Add(iconbitmap);
                                    else
                                    {
                                        bReportError = true;
                                    }
                                    DestroyIcon(large[i]);
                                    if (IntPtr.Zero != small[i])
                                        DestroyIcon(small[i]);
                                }
                            }
                        }
                        else
                        {
                            iconbitmap = ExtractIconBitMap(System.Drawing.Icon.ExtractAssociatedIcon(file));
                            if (null != iconbitmap)
                                ibm.Add(iconbitmap);
                            else
                            {
                                ibm.Add(ExtractIconBitMap(SystemIcons.Error));
                                bReportError = true;
                            }
                        }
                    }
                    else
                    {
                        iconbitmap = ExtractIconBitMap(System.Drawing.Icon.ExtractAssociatedIcon(file));
                        if (null != iconbitmap)
                            ibm.Add(iconbitmap);
                        else
                        {
                            iconbitmap = ExtractIconBitMap(SystemIcons.Error);
                            if (null != iconbitmap)
                                ibm.Add(iconbitmap);
                            bReportError = true;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                iconbitmap = ExtractIconBitMap(SystemIcons.Error);
                if (null != iconbitmap)
                    ibm.Add(iconbitmap);
                lnkio.WriteProgramLog("Icon source anomaly: Failed to load one or more icon(s) from " +
                    file + "System Error: " + e.Message);
                bReportError = false;
            }
            finally
            {
                // RELEASE RESOURCES
                foreach (IntPtr ptr in large)
                    if (ptr != IntPtr.Zero)
                        DestroyIcon(ptr);

                foreach (IntPtr ptr in small)
                    if (ptr != IntPtr.Zero)
                        DestroyIcon(ptr);
                if (null == iconbitmap)
                {
                    iconbitmap = ExtractIconBitMap(SystemIcons.Error);
                    if (null != iconbitmap)
                        ibm.Add(iconbitmap);
                    else
                        bReportError = true;
                }
            }

            if (bReportError )
                lnkio.WriteProgramLog("Icon source anomaly: Failed to load one or more icon(s) from " + file);

            return ibm;
        }
        private static IconBitMap GetBitmapFromFileIcon(string file, int iconIndex)
        {
            IconBitMap ibm = null;
            try
            {
                // if a directory path is used just get the system icon
                if (File.GetAttributes(file).HasFlag(FileAttributes.Directory))
                {
                    // get a 'built in' icon for a folder
                    try
                    {
                        ibm = ExtractIconBitMapFromFile(
                            Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\shell32.dll"), 4);
                    }
                    catch
                    {
                        ibm = null; // ExtractIconBitMap(SystemIcons.Error);
                    }

                }
                else
                {
                    ibm = ExtractICO(file, iconIndex);
                    //ibm = ExtractICO(file, USE_MAIN_ICON);
                }
            }
            catch (Exception e)
            {
                ibm = ExtractIconBitMap(SystemIcons.Error);
                lnkio.WriteProgramLog("Did not get a bitmap source from the file: " + file + " System Error: " + e.Message);
            }

            if (null == ibm)
            {
                ibm = ExtractIconBitMap(SystemIcons.Error);
                lnkio.WriteProgramLog("Failed to load icon for " + file);
            }

            return ibm;
        }

        /// Sources of icons include: 
        /// SystemIcons.Hand
        /// Environment.ExpandEnvironmentVariables(%SystemRoot%\system32\shell32.dll)); 
        /// C:\Windows\System32\imageres.dll 
        /// %SystemRoot%\system32\DDORes.dll

        public static IconBitMap ExtractICO(string file, int idxIcon)
        {
            IconBitMap ibm = null;
            try
            {
                if (USE_MAIN_ICON == idxIcon)
                {
                    //System.Drawing.Icon ico = System.Drawing.Icon.ExtractAssociatedIcon(file);
                    ibm = ExtractIconBitMap(System.Drawing.Icon.ExtractAssociatedIcon(file));
                    //ibm = GetBitmapFromFileIcon(file, USE_MAIN_ICON);
                }
                else
                {
                    ibm = ExtractIconBitMapFromFile(file, idxIcon);
                }
            }
            catch( Exception e )
            {
                if (File.GetAttributes(file).HasFlag(FileAttributes.Directory))
                {
                    // get a 'built in' icon for a folder
                    try
                    {
                        ibm = ExtractIconBitMapFromFile(
                            Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\shell32.dll"), 4);
                    }
                    catch
                    {
                        ibm = null; // ExtractIconBitMap(SystemIcons.Error);
                    }
                }
                else
                {
                    ibm = ExtractIconBitMap(SystemIcons.Error);
                    lnkio.WriteProgramLog("Did not get a bitmap source from the file: " + file + " System Error: " + e.Message);
                }
            }

            if (null == ibm)
            {

                lnkio.WriteProgramLog("Did not get a bitmap source from the file: " + file);
                ibm = ExtractIconBitMap(SystemIcons.Error);
                //throw new Exception("Did not get a bitmap source from the file: " + file);
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
        public Bitmap bitmap { get; set; }
        //public WriteableBitmap writeablebitmap { get; set; }
    }
}
